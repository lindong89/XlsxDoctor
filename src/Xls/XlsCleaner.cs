using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace XlsxDoctor.Xls
{
    /// <summary>解析出来的一条 NAME（定义名称）记录。</summary>
    public sealed class NameRecordInfo
    {
        public int RecordIndex;
        public int Index0;            // 在 NAME 序列中的 0 基序号
        public string Name = "";
        public ushort Grbit;
        public byte ChKey;            // 内置名编号（fBuiltin=1 时有效）
        public int Itab;              // 作用域工作表下标（-1 = 全局）
        public int NameBytes;
        public string RawNameHex = "";   // 名称字段原始字节，用于核对编码
        public int RgceOffset;
        public int Cce;
        public List<TokenRef> Refs = new List<TokenRef>();
        public bool HasErrorToken;    // 定义里含 #REF!/#NAME? 之类错误令牌 → 已失效
        public bool Parsed;
        public string ParseError = "";
        public bool IsSystemName;     // _xlnm.* 内置名

        /// <summary>公式里 ptgName 用的是 1 基索引，这里是它的 1 基值。</summary>
        public int Index1 { get { return Index0 + 1; } }
    }

    /// <summary>
    /// .xls（CFBF + BIFF8）清理引擎。
    ///
    /// 与 .xlsx 最大的不同：BIFF 里公式按**索引**引用定义名称和外部表，不是按名字。
    /// 所以删任何一条都必须同步重编号所有公式里的索引，否则公式会静默指错 ——
    /// 这种错误 Excel 不会报，只会算错数，比直接打不开更危险。
    ///
    /// 因此这里的策略是"先算清楚谁引用了谁，再动手"：
    ///   1. 解析全部 FORMULA 记录的令牌流，收集 ptgName / ptgRef3d 的索引。
    ///      只要有一条解析不出来，整体放弃（宁可不动，不可改错）。
    ///   2. 解析 NAME 定义自身的令牌流（定义之间会互相引用）。
    ///   3. 算出保留集合，生成索引重映射表，原地改写令牌流（长度不变，可直接覆盖）。
    ///   4. 剔除无人引用的记录，重建 BIFF 流，修正 BOUNDSHEET 的绝对偏移。
    ///   5. 重建 CFBF 容器，然后**重新读回校验**，任何一项不过就丢弃输出。
    /// </summary>
    public static class XlsCleaner
    {
        public static CleanReport Run(string srcPath, CleanOptions opt)
        {
            Stopwatch sw = Stopwatch.StartNew();
            if (opt == null) opt = new CleanOptions();

            CleanReport rep = new CleanReport();
            rep.IsXls = true;
            rep.SourcePath = srcPath;
            rep.Conservative = opt.Conservative;

            try
            {
                if (!File.Exists(srcPath))
                {
                    rep.ValidationErrors.Add("文件不存在：" + srcPath);
                    return rep;
                }

                FileInfo fi = new FileInfo(srcPath);
                rep.SourceSize = fi.Length;

                // ---------- 1. 读容器与工作簿流 ----------
                Log(opt, "读取 CFBF 容器…");
                CfbfFile cfbf = CfbfFile.Read(srcPath);
                CfbfEntry wbEntry = cfbf.FindWorkbookEntry();
                if (wbEntry == null)
                {
                    rep.ValidationErrors.Add("CFBF 目录里找不到 Workbook / Book 流，这不像是 Excel 工作簿。");
                    return rep;
                }

                byte[] stream = cfbf.ReadStream(wbEntry);
                rep.XlsStreamBefore = stream.Length;

                Log(opt, "解析 BIFF 记录…");
                BiffFile biff = BiffFile.Parse(stream);
                rep.XlsBiffRecordsBefore = biff.Records.Count;

                // ---------- 2. 解析 NAME 记录 ----------
                Log(opt, "解析定义名称…");
                List<NameRecordInfo> names = ParseNames(biff);
                rep.NameTotal = names.Count;

                for (int i = 0; i < names.Count; i++)
                {
                    NameRecordInfo n = names[i];
                    if (!n.Parsed)
                    {
                        rep.ValidationErrors.Add("第 " + n.Index1.ToString(CultureInfo.InvariantCulture)
                            + " 条定义名称的令牌流解析失败：" + n.ParseError
                            + "（出于安全考虑放弃清理）");
                        return rep;
                    }
                }

                // ---------- 3. 解析公式，收集引用 ----------
                Log(opt, "扫描公式令牌…");
                UsageScan usage = BiffUsage.Scan(biff);
                rep.XlsFormulaParsed = usage.ParseOk;

                if (!usage.AllParsed)
                {
                    rep.ValidationErrors.Add("有 " + usage.ParseFail.ToString(CultureInfo.InvariantCulture)
                        + " 条公式的令牌流无法解析，出于安全考虑放弃清理。");
                    for (int i = 0; i < usage.FailSamples.Count; i++)
                        rep.ValidationErrors.Add("  " + usage.FailSamples[i]);
                    return rep;
                }

                // SHRFMLA（共享公式）里也有令牌流，必须一起扫
                List<FormulaInfo> shr = ScanShrFmla(biff, usage, rep);
                if (rep.ValidationErrors.Count > 0) return rep;

                // ---------- 4. 外部引用表 ----------
                RefTable rt = RefTable.Parse(biff);
                rt.MarkReferenced(usage.ReferencedIxti);

                // ---------- 5. 决定保留 / 删除 ----------
                HashSet<int> keepName = new HashSet<int>();      // 0 基序号
                HashSet<int> keepIxti = new HashSet<int>(usage.ReferencedIxti);

                // 自引用（虚拟）SUPBOOK 代表"本工作簿自己"。指向它的 EXTERNSHEET 条目是三维引用
                // （Sheet1:Sheet3!A1）的基础，属于正常内容而不是外部链接，必须无条件保留 ——
                // 因为图表子流、条件格式里也可能有三维引用，而我们只扫了工作表公式和共享公式。
                int toSelf = 0;
                if (rt.VirtualSupbookIndex >= 0)
                {
                    for (int i = 0; i < rt.ExternSheet.Count; i++)
                    {
                        if (rt.ExternSheet[i].ISupBook != rt.VirtualSupbookIndex) continue;
                        keepIxti.Add(rt.ExternSheet[i].Ixti);
                        toSelf++;
                    }
                }
                rep.XlsExternSheetToSelf = toSelf;

                // 种子 1：被公式引用的名称（ptgName 用的是 1 基索引）
                foreach (int idx1 in usage.ReferencedNameIndices)
                {
                    int i0 = idx1 - 1;
                    if (i0 >= 0 && i0 < names.Count) keepName.Add(i0);
                    // 索引 0 在 1 基口径下不合法；若真的出现，保守起见把第 0 条也留下
                    if (idx1 == 0 && names.Count > 0) keepName.Add(0);
                }

                // 种子 2：系统名（打印区域 / 打印标题 / 筛选区域…）无条件保留。
                // 必须赶在闭包之前加进来：系统名的定义本身也含 ptgRef3d，
                // 晚一步它引用的外部表索引就进不了 keepIxti，改写令牌流时会找不到映射。
                for (int i = 0; i < names.Count; i++)
                    if (names[i].IsSystemName) keepName.Add(i);

                // 传递闭包 (a) 与兜底扫描 (b) 交替迭代，直到不再增长。
                //   (a) 被保留的名称，它定义里引用到的名称/外部表也必须保留；
                //   (b) 凡是会被改写令牌流的记录（公式、共享公式），其引用必须在保留集合里。
                // 交替是因为 (b) 可能引入新的保留项，而新保留项自己又有引用。
                // 这样"引用完整性"就是结构性保证，不再依赖各段代码的先后顺序 ——
                // 之前正是因为顺序写反，导致系统名引用的外部表被误删。
                for (int pass = 0; pass < 16; pass++)
                {
                    bool grew = false;

                    List<int> snapshot = new List<int>(keepName);
                    for (int k = 0; k < snapshot.Count; k++)
                    {
                        NameRecordInfo n = names[snapshot[k]];
                        if (AbsorbRefs(n.Refs, names.Count, keepName, keepIxti)) grew = true;
                    }

                    for (int i = 0; i < usage.Formulas.Count; i++)
                        if (AbsorbRefs(usage.Formulas[i].Refs, names.Count, keepName, keepIxti)) grew = true;
                    for (int i = 0; i < shr.Count; i++)
                        if (AbsorbRefs(shr[i].Refs, names.Count, keepName, keepIxti)) grew = true;

                    if (!grew) break;
                }

                // 分类统计 + 决定删除集合
                HashSet<int> dropName = new HashSet<int>();
                int broken = 0, system = 0, used = 0, unreferenced = 0;
                for (int i = 0; i < names.Count; i++)
                {
                    NameRecordInfo n = names[i];
                    if (n.IsSystemName) { system++; continue; }   // 已在 keepName 里
                    if (keepName.Contains(i)) { used++; continue; }

                    bool isBroken = IsBrokenDefinition(n, rt);
                    if (isBroken) broken++; else unreferenced++;

                    if (opt.Conservative && !isBroken) continue;   // 保守模式留着
                    dropName.Add(i);
                }

                // EXTERNSHEET：保留被引用的
                HashSet<int> dropIxti = new HashSet<int>();
                for (int i = 0; i < rt.ExternSheet.Count; i++)
                {
                    ExternSheetEntry e = rt.ExternSheet[i];
                    if (keepIxti.Contains(e.Ixti)) continue;
                    dropIxti.Add(e.Ixti);
                }

                // SUPBOOK：被保留的 EXTERNSHEET 条目引用到的才留
                HashSet<int> keepSupbook = new HashSet<int>();
                for (int i = 0; i < rt.ExternSheet.Count; i++)
                {
                    ExternSheetEntry e = rt.ExternSheet[i];
                    if (dropIxti.Contains(e.Ixti)) continue;
                    if (e.ISupBook < rt.Supbooks.Count) keepSupbook.Add(e.ISupBook);
                }

                // EXTERNNAME：被 ptgNameX 引用的，或所属 SUPBOOK 被保留的
                HashSet<int> dropExternName = new HashSet<int>();
                int externNameTotal = 0;
                for (int i = 0; i < biff.Records.Count; i++)
                {
                    if (biff.Records[i].Type != 0x0023) continue;
                    externNameTotal++;
                    int owner = OwningSupbook(biff, rt, i);
                    if (owner >= 0 && keepSupbook.Contains(owner)) continue;
                    dropExternName.Add(i);
                }

                rep.NameUsed = used;
                rep.NameSystem = system;
                rep.NameBroken = broken;
                rep.NameUnreferenced = unreferenced;

                rep.LinkTotal = rt.ExternSheet.Count;
                rep.LinkKeptByFormula = rt.ExternSheet.Count - dropIxti.Count;
                rep.LinkDropped = dropIxti.Count;
                rep.XlsVirtualSupbook = rt.VirtualSupbookIndex;
                rep.XlsSupbookTotal = rt.Supbooks.Count;
                rep.XlsSupbookKept = keepSupbook.Count;
                rep.XlsExternSheetTotal = rt.ExternSheet.Count;
                rep.XlsExternSheetKept = rt.ExternSheet.Count - dropIxti.Count;
                rep.XlsExternNameTotal = externNameTotal;
                rep.XlsExternNameDropped = dropExternName.Count;
                rep.XlsFormulaRefs = usage.PtgNameRefs + usage.PtgNameXRefs + usage.ThreeDRefs;

                int sheetCount = biff.CountOf(0x0085);
                rep.SheetCount = sheetCount;

                // 保留的外部工作簿路径，报告里列出来
                for (int i = 0; i < rt.ExternSheet.Count; i++)
                {
                    ExternSheetEntry e = rt.ExternSheet[i];
                    if (dropIxti.Contains(e.Ixti)) continue;
                    if (e.ISupBook < rt.Supbooks.Count)
                    {
                        string p = rt.Supbooks[e.ISupBook].Path;
                        if (!string.IsNullOrEmpty(p) && !rep.KeptLinkTargets.Contains(p))
                            rep.KeptLinkTargets.Add(p);
                    }
                    if (rep.KeptLinkTargets.Count >= 10) break;
                }

                // 有没有东西可删
                bool anythingToDrop = dropName.Count > 0 || dropIxti.Count > 0
                                   || dropExternName.Count > 0 || (rt.Supbooks.Count - keepSupbook.Count) > 0;

                // 顺带检查：EXTERNSHEET 记录重写后会不会超长
                int newExtLen = 2 + (rt.ExternSheet.Count - dropIxti.Count) * 6;
                if (newExtLen > BiffFile.MaxRecordData)
                {
                    rep.ValidationErrors.Add("重写后的 EXTERNSHEET 记录会超过 8224 字节（"
                        + newExtLen.ToString(CultureInfo.InvariantCulture)
                        + "），需要重新做 CONTINUE 分段，本版本暂不支持，放弃清理。");
                    return rep;
                }

                rep.Analyzed = true;

                if (!anythingToDrop)
                {
                    rep.NothingToDo = true;
                    rep.ValidationPassed = true;
                    return rep;
                }

                if (opt.DryRun)
                {
                    rep.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                    return rep;
                }

                // ---------- 6. 生成索引重映射表 ----------
                Log(opt, "生成索引重映射表…");
                int[] nameMap = BuildNameMap(names, dropName);         // 旧 1 基索引 -> 新 1 基索引
                int[] ixtiMap = BuildIxtiMap(rt, dropIxti);            // 旧 ixti -> 新 ixti
                int[] supMap = BuildSupMap(rt, keepSupbook);           // 旧 iSupBook -> 新 iSupBook

                // ---------- 7. 改写令牌流 ----------
                Log(opt, "改写公式索引…");
                HashSet<int> dropRecords = new HashSet<int>();

                for (int i = 0; i < names.Count; i++)
                {
                    if (dropName.Contains(i)) { dropRecords.Add(names[i].RecordIndex); continue; }
                    NameRecordInfo n = names[i];
                    if (n.Refs.Count == 0) continue;
                    byte[] data = (byte[])biff.Records[n.RecordIndex].Data.Clone();
                    byte[] rgce = new byte[n.Cce];
                    Buffer.BlockCopy(data, n.RgceOffset, rgce, 0, n.Cce);
                    BiffFormula.Remap(rgce, n.Refs, nameMap, ixtiMap);
                    Buffer.BlockCopy(rgce, 0, data, n.RgceOffset, n.Cce);
                    biff.Records[n.RecordIndex].NewData = data;
                }

                RewriteFormulas(biff, usage.Formulas, nameMap, ixtiMap);
                RewriteFormulas(biff, shr, nameMap, ixtiMap);

                for (int i = 0; i < biff.Records.Count; i++)
                    if (biff.Records[i].Type == 0x01AE) { /* SUPBOOK 稍后统一处理 */ }

                // 剔除 SUPBOOK / EXTERNNAME / EXTERNSHEET
                for (int i = 0; i < rt.Supbooks.Count; i++)
                    if (!keepSupbook.Contains(i)) dropRecords.Add(rt.Supbooks[i].RecordIndex);
                foreach (int i in dropExternName) dropRecords.Add(i);

                // 重写 EXTERNSHEET（只留被引用的条目，并更新 iSupBook）
                int esIndex = FindExternSheetRecord(biff);
                if (esIndex >= 0 && dropIxti.Count > 0)
                {
                    List<ExternSheetEntry> kept = new List<ExternSheetEntry>();
                    for (int i = 0; i < rt.ExternSheet.Count; i++)
                        if (!dropIxti.Contains(rt.ExternSheet[i].Ixti)) kept.Add(rt.ExternSheet[i]);

                    byte[] nd = new byte[2 + kept.Count * 6];
                    nd[0] = (byte)(kept.Count & 0xFF);
                    nd[1] = (byte)((kept.Count >> 8) & 0xFF);
                    for (int k = 0; k < kept.Count; k++)
                    {
                        int ns = (kept[k].ISupBook < supMap.Length) ? supMap[kept[k].ISupBook] : -1;
                        if (ns < 0) throw new XlsFormatException("内部错误：保留的 EXTERNSHEET 条目指向了被删除的 SUPBOOK");
                        int off = 2 + k * 6;
                        nd[off] = (byte)(ns & 0xFF); nd[off + 1] = (byte)((ns >> 8) & 0xFF);
                        nd[off + 2] = (byte)(kept[k].ItabFirst & 0xFF); nd[off + 3] = (byte)((kept[k].ItabFirst >> 8) & 0xFF);
                        nd[off + 4] = (byte)(kept[k].ItabLast & 0xFF); nd[off + 5] = (byte)((kept[k].ItabLast >> 8) & 0xFF);
                    }
                    biff.Records[esIndex].NewData = nd;
                }

                // ---------- 8. 重建 BIFF 流，并修正绝对偏移 ----------
                Log(opt, "重建工作簿流…");
                Dictionary<int, int> offsetMap;
                byte[] rebuilt = biff.Rebuild(dropRecords, out offsetMap);

                // BOUNDSHEET 的 lbPlyPos 是各工作表 BOF 在流中的绝对偏移，删字节后必须跟着挪
                int fixedBound = FixupBoundsheets(biff, dropRecords, offsetMap);
                int fixedExtSst = FixupExtSst(biff, dropRecords, offsetMap);

                // 偏移修好后再重建一次，得到最终字节
                rebuilt = biff.Rebuild(dropRecords, out offsetMap);
                rep.XlsStreamAfter = rebuilt.Length;

                // ---------- 9. 写文件 ----------
                string outPath = ResolveOutputPath(srcPath, opt);

                // 防呆：输出路径与源文件相同时必须拒绝，否则"清理"会变成"原地破坏"。
                // 要覆盖原文件得显式用「就地覆盖」选项，那条路会先自动备份。
                if (!opt.InPlace && string.Equals(Path.GetFullPath(outPath), Path.GetFullPath(srcPath),
                                                  StringComparison.OrdinalIgnoreCase))
                {
                    rep.ValidationErrors.Add("输出路径不能与源文件相同。若要覆盖原文件，请使用「就地覆盖」选项（会自动备份）。");
                    return rep;
                }

                string backupPath = null;
                if (opt.InPlace)
                {
                    backupPath = srcPath + ".bak";
                    if (File.Exists(backupPath)) File.Delete(backupPath);
                    File.Copy(srcPath, backupPath, true);
                    outPath = srcPath;
                }

                Log(opt, "重建 CFBF 容器…");
                Dictionary<int, byte[]> repl = new Dictionary<int, byte[]>();
                repl[wbEntry.Index] = rebuilt;
                CfbfWriter.Write(outPath, cfbf, repl);

                rep.OutputPath = outPath;
                rep.BackupPath = backupPath;
                rep.WroteFile = true;
                rep.OutputSize = new FileInfo(outPath).Length;
                rep.DroppedParts = dropRecords.Count;

                BiffFile outBiff = BiffFile.Parse(rebuilt);
                rep.XlsBiffRecordsAfter = outBiff.Records.Count;

                // ---------- 10. 读回校验 ----------
                Log(opt, "校验输出…");
                ValidateOutput(outPath, rep, usage.FormulaRecords, names.Count - dropName.Count,
                               rt.ExternSheet.Count - dropIxti.Count);

                if (rep.ValidationErrors.Count > 0 && rep.WroteFile && !opt.InPlace)
                {
                    try { File.Delete(outPath); } catch (Exception) { }
                    rep.WroteFile = false;
                }

                rep.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                return rep;
            }
            catch (XlsFormatException ex)
            {
                rep.ValidationErrors.Add("格式错误：" + ex.Message);
                return rep;
            }
            catch (Exception ex)
            {
                rep.ValidationErrors.Add("未预期的错误：" + ex.GetType().Name + " — " + ex.Message);
                return rep;
            }
        }

        // ================================================================== NAME 解析

        /// <summary>
        /// 解析 NAME(0x0018) 记录。布局：
        ///   grbit(2) chKey(1) cch(1) cce(2) ixals(2) itab(2)
        ///   cchCustMenu(1) cchDescription(1) cchHelptopic(1) cchStatusbar(1)   = 14 字节头
        ///   名称字符串：1 字节选项位 + (cch 或 2*cch) 字节
        ///   rgce：cce 字节
        /// </summary>
        public static List<NameRecordInfo> ParseNames(BiffFile biff)
        {
            List<NameRecordInfo> list = new List<NameRecordInfo>();

            for (int i = 0; i < biff.Records.Count; i++)
            {
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x0018) continue;

                NameRecordInfo n = new NameRecordInfo();
                n.RecordIndex = i;
                n.Index0 = list.Count;
                list.Add(n);

                byte[] d = r.Data;
                if (d.Length < 15) { n.ParseError = "记录长度不足（" + d.Length.ToString(CultureInfo.InvariantCulture) + " 字节）"; continue; }

                n.Grbit = BitConverter.ToUInt16(d, 0);
                n.ChKey = d[2];
                int cch = d[3];
                n.Cce = BitConverter.ToUInt16(d, 4);
                n.Itab = BitConverter.ToInt16(d, 8);

                int nameBytes = cch;
                if (d.Length > 14)
                {
                    int optByte = d[14];
                    nameBytes = ((optByte & 0x01) != 0) ? cch * 2 : cch;
                    if ((optByte & 0x01) != 0 && cch > 0 && d.Length >= 15 + nameBytes)
                        n.Name = Encoding.Unicode.GetString(d, 15, nameBytes);
                    else if (cch > 0 && d.Length >= 15 + nameBytes)
                    {
                        char[] cs = new char[nameBytes];
                        for (int k = 0; k < nameBytes; k++) cs[k] = (char)d[15 + k];
                        n.Name = new string(cs);
                    }
                }

                n.NameBytes = nameBytes;

                int hx = Math.Min(24, Math.Max(0, d.Length - 15));
                for (int q = 0; q < hx; q++)
                    n.RawNameHex += d[15 + q].ToString("X2", CultureInfo.InvariantCulture) + " ";

                n.IsSystemName = IsBuiltinName(n.Name, n.Grbit, n.ChKey);

                n.RgceOffset = 14 + 1 + nameBytes;
                if (n.RgceOffset + n.Cce > d.Length)
                {
                    n.ParseError = string.Format(CultureInfo.InvariantCulture,
                        "rgce 越界：偏移 {0} + 长度 {1} > 记录长 {2}", n.RgceOffset, n.Cce, d.Length);
                    continue;
                }

                byte[] rgce = new byte[n.Cce];
                Buffer.BlockCopy(d, n.RgceOffset, rgce, 0, n.Cce);

                string err;
                if (!BiffFormula.TryScan(rgce, n.Cce, n.Refs, out err))
                {
                    n.ParseError = err;
                    continue;
                }

                // 定义里出现错误令牌 → 这个名字已经失效
                n.HasErrorToken = ContainsErrorToken(rgce, n.Cce);
                n.Parsed = true;
            }

            return list;
        }

        /// <summary>
        /// 判断是不是内置（系统）名称。
        ///
        /// BIFF8 里有两种表示法，必须都认：
        ///   1. 名称字符串直接以 "_xlnm." 开头（这是 MS-XLS 规定的内置名写法）
        ///   2. grbit 的 bit5（0x20）置位表示 fBuiltin=1，此时名称可能是本地化文本，
        ///      例如中文 Excel 会写成 "Print_Area" 甚至 "工作表1!Print_Area"
        /// 内置名（打印区域、打印标题、筛选区域等）是 Excel 自己维护的，删掉会导致
        /// 打印设置丢失，所以必须无条件保留。
        /// </summary>
        public static bool IsBuiltinName(string name, ushort grbit, byte chKey)
        {
            if (name != null && name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) return true;

            // fBuiltin 位
            if ((grbit & 0x20) != 0) return true;

            // 兜底：按已知的内置名清单比对（去掉 "表名!" 前缀后比较）
            if (string.IsNullOrEmpty(name)) return false;
            string bare = name;
            int bang = bare.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < bare.Length) bare = bare.Substring(bang + 1);

            string[] builtins = new string[]
            {
                "Consolidate_Area", "Auto_Open", "Auto_Close", "Extract", "Database",
                "Criteria", "Print_Area", "Print_Titles", "Recorder", "Data_Form",
                "Auto_Activate", "Auto_Deactivate", "Sheet_Title", "_FilterDatabase",
                "FilterDatabase"
            };
            for (int i = 0; i < builtins.Length; i++)
                if (string.Equals(bare, builtins[i], StringComparison.OrdinalIgnoreCase)) return true;

            return false;
        }

        /// <summary>
        /// 判定一条定义名称是否已经失效（界面上显示为 #REF!）。
        ///
        /// 光看错误令牌是不够的：工作表被删除后，Excel 并不会在 BIFF 里写入 ptgRefErr，
        /// 而是留下一个 ptgRef3d，其 ixti 指向的 EXTERNSHEET 条目已经超出工作表范围。
        /// 这种"隐式失效"必须靠核对 ixti -> EXTERNSHEET -> SUPBOOK.ctab 才能发现。
        /// </summary>
        static bool IsBrokenDefinition(NameRecordInfo n, RefTable rt)
        {
            if (n.HasErrorToken) return true;

            // 工作表作用域的名称，itab 是 1 基下标；0 表示全局。
            if (n.Itab > 0)
            {
                int sheetCount = -1;
                if (rt.VirtualSupbookIndex >= 0 && rt.VirtualSupbookIndex < rt.Supbooks.Count)
                    sheetCount = rt.Supbooks[rt.VirtualSupbookIndex].Ctab;
                if (sheetCount > 0 && n.Itab > sheetCount) return true;
            }

            for (int i = 0; i < n.Refs.Count; i++)
            {
                TokenRef tr = n.Refs[i];
                if (tr.Ixti < 0) continue;
                if (tr.Ixti >= rt.ExternSheet.Count) return true;

                ExternSheetEntry e = rt.ExternSheet[tr.Ixti];
                if (e.ISupBook < 0 || e.ISupBook >= rt.Supbooks.Count) return true;

                SupbookInfo sb = rt.Supbooks[e.ISupBook];
                if (sb.Ctab > 0 && (e.ItabFirst >= sb.Ctab || e.ItabLast >= sb.Ctab)) return true;
            }

            return false;
        }

        /// <summary>rgce 里是否含错误令牌 —— 表示定义已经失效（例如指向了被删掉的工作表）。</summary>
        static bool ContainsErrorToken(byte[] rgce, int cce)
        {
            int pos = 0;
            int guard = 0;
            while (pos < cce && guard++ < 200000)
            {
                string tn;
                int len = BiffFormula.LengthOfToken(rgce, pos, out tn);
                if (len <= 0) return false;

                byte ptg = rgce[pos];
                if (ptg >= 0x20)
                {
                    // 类化令牌：ptgRefErr / ptgAreaErr / ptgRefErr3d / ptgAreaErr3d
                    int id = ptg & 0x1F;
                    if (id == 0x0A || id == 0x0B || id == 0x1C || id == 0x1D) return true;
                }
                else if (ptg == 0x1C)
                {
                    // 非类化 ptgErr（字节 0x1C 且 < 0x20）：定义本身就是一个错误字面量，
                    // 例如 =#REF! —— 这是名称已失效最直接的表示。
                    return true;
                }
                pos += len;
            }
            return false;
        }

        // ================================================================== SHRFMLA

        /// <summary>SHRFMLA(0x04BC) 共享公式：refU(6) + cce(2) + rgce。rgce 从偏移 8 开始。</summary>
        static List<FormulaInfo> ScanShrFmla(BiffFile biff, UsageScan usage, CleanReport rep)
        {
            List<FormulaInfo> list = new List<FormulaInfo>();

            for (int i = 0; i < biff.Records.Count; i++)
            {
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x04BC) continue;
                if (r.Data.Length < 8) continue;

                int cce = BitConverter.ToUInt16(r.Data, 6);
                if (8 + cce > r.Data.Length) continue;

                byte[] rgce = new byte[cce];
                Buffer.BlockCopy(r.Data, 8, rgce, 0, cce);

                List<TokenRef> refs = new List<TokenRef>();
                string err;
                if (!BiffFormula.TryScan(rgce, cce, refs, out err))
                {
                    rep.ValidationErrors.Add("共享公式记录 #" + i.ToString(CultureInfo.InvariantCulture)
                        + " 的令牌流无法解析：" + err + "（出于安全考虑放弃清理）");
                    return list;
                }

                FormulaInfo fi = new FormulaInfo();
                fi.RecordIndex = i;
                fi.RgceOffset = 8;
                fi.Cce = cce;
                fi.Refs = refs;
                list.Add(fi);

                for (int k = 0; k < refs.Count; k++)
                {
                    if (refs[k].NameIndex >= 0) usage.ReferencedNameIndices.Add(refs[k].NameIndex);
                    if (refs[k].Ixti >= 0) usage.ReferencedIxti.Add(refs[k].Ixti);
                }
            }

            return list;
        }

        static void RewriteFormulas(BiffFile biff, List<FormulaInfo> list, int[] nameMap, int[] ixtiMap)
        {
            for (int i = 0; i < list.Count; i++)
            {
                FormulaInfo fi = list[i];
                if (fi.Refs.Count == 0) continue;

                BiffRecord r = biff.Records[fi.RecordIndex];
                byte[] data = (byte[])r.Data.Clone();
                byte[] rgce = new byte[fi.Cce];
                Buffer.BlockCopy(data, fi.RgceOffset, rgce, 0, fi.Cce);
                BiffFormula.Remap(rgce, fi.Refs, nameMap, ixtiMap);
                Buffer.BlockCopy(rgce, 0, data, fi.RgceOffset, fi.Cce);
                r.NewData = data;
            }
        }

        // ================================================================== 映射表

        /// <summary>
        /// 把一批令牌引用并入保留集合；返回是否有新增项。
        /// ptgName 的 nameindex 是 1 基的，所以要减 1 才对应 NAME 记录的 0 基序号。
        /// </summary>
        static bool AbsorbRefs(List<TokenRef> refs, int nameCount, HashSet<int> keepName, HashSet<int> keepIxti)
        {
            bool grew = false;
            for (int r = 0; r < refs.Count; r++)
            {
                TokenRef tr = refs[r];
                if (tr.Ixti >= 0 && keepIxti.Add(tr.Ixti)) grew = true;
                if (tr.NameIndex >= 0)
                {
                    int i0 = tr.NameIndex - 1;
                    if (i0 >= 0 && i0 < nameCount && keepName.Add(i0)) grew = true;
                    if (tr.NameIndex == 0 && nameCount > 0 && keepName.Add(0)) grew = true;
                }
            }
            return grew;
        }

        static int[] BuildNameMap(List<NameRecordInfo> names, HashSet<int> drop)
        {
            // ptgName 用 1 基索引，映射表下标也是 1 基；0 号位保留给"特殊值"
            int[] map = new int[names.Count + 2];
            for (int i = 0; i < map.Length; i++) map[i] = -1;

            int next = 1;
            for (int i = 0; i < names.Count; i++)
                if (!drop.Contains(i)) map[i + 1] = next++;
            return map;
        }

        static int[] BuildIxtiMap(RefTable rt, HashSet<int> drop)
        {
            int[] map = new int[rt.ExternSheet.Count];
            for (int i = 0; i < map.Length; i++) map[i] = -1;

            int next = 0;
            for (int i = 0; i < rt.ExternSheet.Count; i++)
                if (!drop.Contains(i)) map[i] = next++;
            return map;
        }

        static int[] BuildSupMap(RefTable rt, HashSet<int> keep)
        {
            int[] map = new int[rt.Supbooks.Count];
            for (int i = 0; i < map.Length; i++) map[i] = -1;

            int next = 0;
            for (int i = 0; i < rt.Supbooks.Count; i++)
                if (keep.Contains(i)) map[i] = next++;
            return map;
        }

        /// <summary>EXTERNNAME 归属的 SUPBOOK：它前面最近的一条 SUPBOOK。</summary>
        static int OwningSupbook(BiffFile biff, RefTable rt, int recordIndex)
        {
            int best = -1;
            for (int i = 0; i < rt.Supbooks.Count; i++)
            {
                if (rt.Supbooks[i].RecordIndex < recordIndex) best = i;
                else break;
            }
            return best;
        }

        static int FindExternSheetRecord(BiffFile biff)
        {
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].Type == 0x0017) return i;
            return -1;
        }

        // ================================================================== 偏移修正

        /// <summary>BOUNDSHEET(0x0085) 前 4 字节是工作表 BOF 在流中的绝对偏移，删字节后必须重算。</summary>
        static int FixupBoundsheets(BiffFile biff, HashSet<int> drop, Dictionary<int, int> offsetMap)
        {
            int fixedCount = 0;
            for (int i = 0; i < biff.Records.Count; i++)
            {
                if (drop.Contains(i)) continue;
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x0085) continue;
                if (r.Data.Length < 4) continue;

                uint oldPos = BitConverter.ToUInt32(r.Data, 0);
                int newPos;
                if (!offsetMap.TryGetValue((int)oldPos, out newPos)) continue;

                byte[] d = (r.NewData != null) ? (byte[])r.NewData.Clone() : (byte[])r.Data.Clone();
                d[0] = (byte)(newPos & 0xFF);
                d[1] = (byte)((newPos >> 8) & 0xFF);
                d[2] = (byte)((newPos >> 16) & 0xFF);
                d[3] = (byte)((newPos >> 24) & 0xFF);
                r.NewData = d;
                fixedCount++;
            }
            return fixedCount;
        }

        /// <summary>EXTSST(0x00FF) 里存的是 SST 分桶的绝对流偏移，也要跟着挪。</summary>
        static int FixupExtSst(BiffFile biff, HashSet<int> drop, Dictionary<int, int> offsetMap)
        {
            int fixedCount = 0;
            for (int i = 0; i < biff.Records.Count; i++)
            {
                if (drop.Contains(i)) continue;
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x00FF) continue;
                if (r.Data.Length < 4) continue;

                int cstTotal = BitConverter.ToUInt16(r.Data, 0);
                int cstUnique = BitConverter.ToUInt16(r.Data, 2);
                int entries = Math.Min(cstTotal, cstUnique);
                if (entries <= 0) continue;
                if (4 + entries * 8 > r.Data.Length) continue;

                byte[] d = (byte[])r.Data.Clone();
                bool changed = false;
                for (int k = 0; k < entries; k++)
                {
                    int off = 4 + k * 8;
                    int oldIb = (int)BitConverter.ToUInt32(d, off);
                    int newIb;
                    if (!offsetMap.TryGetValue(oldIb, out newIb)) continue;
                    d[off] = (byte)(newIb & 0xFF);
                    d[off + 1] = (byte)((newIb >> 8) & 0xFF);
                    d[off + 2] = (byte)((newIb >> 16) & 0xFF);
                    d[off + 3] = (byte)((newIb >> 24) & 0xFF);
                    changed = true;
                }
                if (changed) { r.NewData = d; fixedCount++; }
            }
            return fixedCount;
        }

        // ================================================================== 校验

        static void ValidateOutput(string outPath, CleanReport rep, int expectFormulas,
                                   int expectNames, int expectIxti)
        {
            try
            {
                CfbfFile cf = CfbfFile.Read(outPath);
                CfbfEntry wb = cf.FindWorkbookEntry();
                if (wb == null)
                {
                    rep.ValidationErrors.Add("校验：输出文件里找不到 Workbook 流。");
                    return;
                }

                byte[] st = cf.ReadStream(wb);
                BiffFile bf = BiffFile.Parse(st);

                int formulas = bf.CountOf(0x0006);
                if (formulas != expectFormulas)
                    rep.ValidationErrors.Add("校验：FORMULA 记录数从 " + expectFormulas.ToString(CultureInfo.InvariantCulture)
                        + " 变成了 " + formulas.ToString(CultureInfo.InvariantCulture) + "。");

                int names = bf.CountOf(0x0018);
                if (names != expectNames)
                    rep.ValidationErrors.Add("校验：NAME 记录数从 " + expectNames.ToString(CultureInfo.InvariantCulture)
                        + " 变成了 " + names.ToString(CultureInfo.InvariantCulture) + "。");

                // 所有公式必须仍能解析，且索引不越界
                UsageScan us = BiffUsage.Scan(bf);
                if (!us.AllParsed)
                    rep.ValidationErrors.Add("校验：输出里有 " + us.ParseFail.ToString(CultureInfo.InvariantCulture)
                        + " 条公式无法解析。");

                foreach (int v in us.ReferencedNameIndices)
                {
                    if (v < 0 || v > names)
                    {
                        rep.ValidationErrors.Add("校验：输出里出现越界的名称索引 " + v.ToString(CultureInfo.InvariantCulture)
                            + "（当前名称数 " + names.ToString(CultureInfo.InvariantCulture) + "）。");
                        break;
                    }
                }

                int esIdx = FindExternSheetRecord(bf);
                if (esIdx >= 0)
                {
                    int cXti = BitConverter.ToUInt16(bf.Records[esIdx].Data, 0);
                    if (cXti != expectIxti)
                        rep.ValidationErrors.Add("校验：EXTERNSHEET 条目数从 " + expectIxti.ToString(CultureInfo.InvariantCulture)
                            + " 变成了 " + cXti.ToString(CultureInfo.InvariantCulture) + "。");
                    foreach (int v in us.ReferencedIxti)
                    {
                        if (v >= cXti)
                        {
                            rep.ValidationErrors.Add("校验：输出里出现越界的 ixti " + v.ToString(CultureInfo.InvariantCulture)
                                + "（当前 EXTERNSHEET 条目数 " + cXti.ToString(CultureInfo.InvariantCulture) + "）。");
                            break;
                        }
                    }
                }

                // BOUNDSHEET 的偏移必须指向真正的 BOF 记录
                HashSet<int> starts = new HashSet<int>();
                for (int i = 0; i < bf.Records.Count; i++) starts.Add(bf.Records[i].Start);

                for (int i = 0; i < bf.Records.Count; i++)
                {
                    BiffRecord r = bf.Records[i];
                    if (r.Type != 0x0085 || r.Data.Length < 4) continue;
                    int pos = (int)BitConverter.ToUInt32(r.Data, 0);
                    if (!starts.Contains(pos))
                    {
                        rep.ValidationErrors.Add("校验：BOUNDSHEET 指向的偏移 " + pos.ToString(CultureInfo.InvariantCulture)
                            + " 不是任何记录的开始位置。");
                        break;
                    }
                    if (pos + 4 > st.Length || BitConverter.ToUInt16(st, pos) != 0x0809)
                    {
                        rep.ValidationErrors.Add("校验：BOUNDSHEET 指向的偏移 " + pos.ToString(CultureInfo.InvariantCulture)
                            + " 处不是 BOF 记录。");
                        break;
                    }
                }

                if (rep.ValidationErrors.Count == 0) rep.ValidationPassed = true;
            }
            catch (Exception ex)
            {
                rep.ValidationErrors.Add("校验时出错：" + ex.GetType().Name + " — " + ex.Message);
            }
        }

        // ================================================================== 杂项

        static string ResolveOutputPath(string srcPath, CleanOptions opt)
        {
            if (!string.IsNullOrEmpty(opt.OutputPath)) return opt.OutputPath;
            string dir = Path.GetDirectoryName(srcPath);
            string baseName = Path.GetFileNameWithoutExtension(srcPath);
            string ext = Path.GetExtension(srcPath);
            return Path.Combine(dir, baseName + " - 已清理" + ext);
        }

        static void Log(CleanOptions opt, string msg)
        {
            if (opt != null && opt.Log != null) opt.Log(msg);
        }
    }
}
