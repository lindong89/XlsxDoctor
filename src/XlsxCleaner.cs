using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxDoctor
{
    internal sealed class DefinedNameInfo
    {
        public string Raw;
        public string Name;
        public string LocalSheetId;
        public string Value;
        public bool Broken;
    }

    internal sealed class ExternalRefInfo
    {
        public string Raw;
        public string RId;
        public bool Keep;
    }

    internal sealed class RelationshipInfo
    {
        public string Raw;
        public string Id;
        public string Type;
        public string Target;
    }

    /// <summary>
    /// xlsx / xlsm 垃圾清理引擎。
    ///
    /// 只动 xl/workbook.xml、xl/_rels/workbook.xml.rels、[Content_Types].xml 和 xl/externalLinks/*，
    /// 其余部件（工作表、图表、绘图、样式、共享字符串…）逐字节原样保留。
    ///
    /// 删除策略：
    ///   1. 值为 #REF!/#N/A 等错误的定义名称                -> 删
    ///   2. 没有任何公式/图表/条件格式/数据验证引用的定义名称 -> 删
    ///   3. 有效的 _xlnm.* 系统名称（打印区域/打印标题/筛选） -> 保留
    ///   4. 被引用的定义名称                                -> 保留
    ///   5. 没有被任何公式引用的外部链接                     -> 删（连同 .rels 与 Content_Types 记录）
    ///   6. 被公式引用的外部链接                            -> 保留
    /// </summary>
    public sealed class XlsxCleaner
    {
        // ---------------- 正则 ----------------
        static readonly Regex FormulaRx = new Regex(
            "<(?:[A-Za-z]+:)?(?:f|formula|formula1|formula2)\\b[^>]*>([^<]*)</(?:[A-Za-z]+:)?(?:f|formula|formula1|formula2)>",
            RegexOptions.Compiled);
        static readonly Regex DefinedNameRx = new Regex("<definedName\\b[^>]*?(?:/>|>[^<]*</definedName>)", RegexOptions.Compiled);
        static readonly Regex DnAttrsRx = new Regex("^<definedName\\b([^>]*?)/?>", RegexOptions.Compiled);
        static readonly Regex DnValueRx = new Regex(">([^<]*)</definedName>$", RegexOptions.Compiled);
        static readonly Regex BrokenRx = new Regex("#(REF!|N/A|NAME\\?|VALUE!|DIV/0!|NULL!|NUM!|GETTING_DATA)", RegexOptions.Compiled);
        static readonly Regex ExtRefRx = new Regex("<externalReference\\b([^>]*)/>", RegexOptions.Compiled);
        static readonly Regex RelationshipRx = new Regex("<Relationship\\b([^>]*)/>", RegexOptions.Compiled);
        static readonly Regex OverrideRx = new Regex("<Override\\b([^>]*)/>", RegexOptions.Compiled);
        // Excel names may contain Unicode letters outside the CJK block (for
        // example Japanese, Cyrillic and Greek).  Restricting this to ASCII +
        // Chinese can make a live name look unused and delete it.
        static readonly Regex IdentRx = new Regex(@"[\p{L}\p{Nl}_\\][\p{L}\p{Nl}\p{Nd}_.\\]*", RegexOptions.Compiled);
        static readonly Regex StringLitRx = new Regex("\"[^\"]*\"", RegexOptions.Compiled);
        static readonly Regex SingleQuotedRx = new Regex("'[^']*'", RegexOptions.Compiled);
        static readonly Regex CellTailRx = new Regex("^\\$?\\d+", RegexOptions.Compiled);
        static readonly Regex FuncTailRx = new Regex("^\\s*\\(", RegexOptions.Compiled);
        static readonly Regex LambdaRx = new Regex("^\\s*(_xlfn\\.)?LAMBDA\\s*\\(", RegexOptions.Compiled);
        static readonly Regex LinkRefRx = new Regex("\\[(\\d+)\\]", RegexOptions.Compiled);
        static readonly Regex PartRx = new Regex("^xl/(worksheets|charts|drawings|dialogsheets|macrosheets)/[^/]+\\.xml$", RegexOptions.Compiled);
        static readonly Regex ExtLinkPartRx = new Regex("^xl/externalLinks/[^/]+\\.xml$", RegexOptions.Compiled);
        static readonly Regex EmptyDnRx = new Regex("<definedNames>\\s*</definedNames>", RegexOptions.Compiled);
        static readonly Regex EmptyErRx = new Regex("<externalReferences>\\s*</externalReferences>", RegexOptions.Compiled);
        static readonly Regex SheetRx = new Regex("<sheet ", RegexOptions.Compiled);

        const string WorkbookPart = "xl/workbook.xml";
        const string WorkbookRelsPart = "xl/_rels/workbook.xml.rels";
        const string ContentTypesPart = "[Content_Types].xml";

        public CleanReport Run(string srcPath, CleanOptions opt)
        {
            if (opt == null) opt = new CleanOptions();
            Stopwatch sw = Stopwatch.StartNew();
            CleanReport report = new CleanReport();
            report.Conservative = opt.Conservative;

            string fullSrc = Path.GetFullPath(srcPath);
            if (!File.Exists(fullSrc)) throw new FileNotFoundException("源文件不存在: " + fullSrc, fullSrc);

            string ext = Path.GetExtension(fullSrc).ToLowerInvariant();

            // 按**文件内容**判断，不看扩展名 —— 把 .xlsx 改名成 .xls 是很常见的操作，
            // Excel 自己也是按内容识别的，照样能打开。只看扩展名会白白拒绝掉能处理的文件。
            FileKind kind = FormatSniffer.Sniff(fullSrc);

            if (kind == FileKind.Cfbf)
                throw new NotSupportedException(
                    "这是 .xls 文件（CFBF/OLE2 复合文档），暂不支持。\r\n"
                    + "本工具目前只处理 OOXML 包：.xlsx / .xlsm / .xltx / .xltm。\r\n"
                    + "想清理的话，可以先用 Excel 打开它，「另存为」成 .xlsx，再用本工具清理。");

            if (kind != FileKind.Ooxml)
                throw new NotSupportedException(
                    "无法识别的文件格式（扩展名 " + ext + "）：文件头既不是 OOXML 的 ZIP 签名（PK），"
                    + "也不是 .xls 的 CFBF 签名（D0CF11E0）。请确认这是不是一个 Excel 工作簿。");

            // 扩展名和内容不符时提醒一句，但照常处理
            report.AddWarning(FormatSniffer.ExtensionMismatch(fullSrc, kind));

            report.SourcePath = fullSrc;
            report.SourceSize = new FileInfo(fullSrc).Length;

            string dst;
            if (opt.InPlace) dst = fullSrc;
            else if (!string.IsNullOrEmpty(opt.OutputPath)) dst = Path.GetFullPath(opt.OutputPath);
            else dst = Path.Combine(Path.GetDirectoryName(fullSrc),
                     Path.GetFileNameWithoutExtension(fullSrc) + " - 已清理" + ext);

            if (!opt.DryRun && !opt.InPlace && string.Equals(dst, fullSrc, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("输出路径不能与源文件相同。若要覆盖原文件，请使用「就地覆盖」选项（会自动备份）。");

            report.OutputPath = dst;

            // ---------------- 打开源包 ----------------
            Log(opt, "打开源文件…");
            FileStream srcFs = File.Open(fullSrc, FileMode.Open, FileAccess.Read, FileShare.Read);
            ZipArchive srcZip;
            try { srcZip = new ZipArchive(srcFs, ZipArchiveMode.Read); }
            catch (Exception ex)
            {
                srcFs.Dispose();
                throw new InvalidDataException("无法作为 zip/xlsx 打开（文件可能已损坏，或根本不是 OOXML 格式）: " + fullSrc, ex);
            }

            string workPath = null;
            string tempPath = null;   // 仅「就地覆盖」使用的临时文件，收尾时只删它
            try
            {
                List<ZipArchiveEntry> srcEntries = new List<ZipArchiveEntry>();
                Dictionary<string, ZipArchiveEntry> map = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
                foreach (ZipArchiveEntry e in srcZip.Entries)
                {
                    srcEntries.Add(e);
                    if (!map.ContainsKey(e.FullName)) map[e.FullName] = e;
                }

                if (!map.ContainsKey(WorkbookPart))
                    throw new InvalidDataException("不是 OOXML 表格文件（缺少 xl/workbook.xml）: " + fullSrc);

                TextPart wbPart = Util.DecodePart(WorkbookPart, Util.ReadAllBytes(map[WorkbookPart]));
                TextPart relsPart = map.ContainsKey(WorkbookRelsPart)
                    ? Util.DecodePart(WorkbookRelsPart, Util.ReadAllBytes(map[WorkbookRelsPart])) : null;
                TextPart ctPart = map.ContainsKey(ContentTypesPart)
                    ? Util.DecodePart(ContentTypesPart, Util.ReadAllBytes(map[ContentTypesPart])) : null;

                string wb = wbPart.Text;
                string relsText = relsPart == null ? "" : relsPart.Text;
                string ctText = ctPart == null ? "" : ctPart.Text;

                report.WorkbookXmlBefore = wb.Length;
                report.SheetCount = SheetRx.Matches(wb).Count;

                // ---------------- 收集公式文本 ----------------
                Log(opt, "扫描公式（工作表 / 图表 / 条件格式 / 数据验证）…");
                List<string> formulaTexts = new List<string>();
                for (int i = 0; i < srcEntries.Count; i++)
                {
                    ZipArchiveEntry f = srcEntries[i];
                    if (f.Length == 0) continue;
                    if (!PartRx.IsMatch(f.FullName)) continue;
                    string xml = Util.ReadAllText(f);
                    foreach (Match m in FormulaRx.Matches(xml))
                        formulaTexts.Add(m.Groups[1].Value);
                }
                report.FormulaTextCount = formulaTexts.Count;

                // ---------------- 解析定义名称 ----------------
                Log(opt, "解析定义名称…");
                List<DefinedNameInfo> definedNames = new List<DefinedNameInfo>();
                foreach (Match m in DefinedNameRx.Matches(wb))
                {
                    string raw = m.Value;
                    Match am = DnAttrsRx.Match(raw);
                    Match vm = DnValueRx.Match(raw);
                    DefinedNameInfo d = new DefinedNameInfo();
                    d.Raw = raw;
                    d.Name = am.Success ? Util.GetAttr(am.Groups[1].Value, "name") : "";
                    d.LocalSheetId = am.Success ? Util.GetAttr(am.Groups[1].Value, "localSheetId") : "";
                    d.Value = vm.Success ? vm.Groups[1].Value : "";
                    d.Broken = BrokenRx.IsMatch(d.Value);
                    definedNames.Add(d);
                }

                HashSet<string> nameLookup = new HashSet<string>(StringComparer.Ordinal);
                HashSet<string> lambdaNames = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < definedNames.Count; i++)
                {
                    DefinedNameInfo d = definedNames[i];
                    string up = d.Name.ToUpperInvariant();
                    nameLookup.Add(up);
                    // LAMBDA 定义的名称在公式里是当函数调用的（=MyFn(1)），必须算作「被引用」
                    if (LambdaRx.IsMatch(d.Value)) lambdaNames.Add(up);
                }

                // ---------------- 哪些名称真的被引用了 ----------------
                HashSet<string> usedNames = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < formulaTexts.Count; i++)
                    CollectUsedNames(formulaTexts[i], nameLookup, lambdaNames, usedNames);

                // A defined name can depend on another defined name.  A name
                // used by a worksheet formula therefore keeps its complete
                // dependency chain; otherwise deleting an intermediate name
                // silently turns the surviving name into #NAME?.
                bool added;
                do
                {
                    added = false;
                    for (int i = 0; i < definedNames.Count; i++)
                    {
                        DefinedNameInfo d = definedNames[i];
                        if (!usedNames.Contains(d.Name.ToUpperInvariant())) continue;
                        int before = usedNames.Count;
                        CollectUsedNames(d.Value, nameLookup, lambdaNames, usedNames);
                        if (usedNames.Count != before) added = true;
                    }
                } while (added);

                // ---------------- 分类 ----------------
                List<DefinedNameInfo> keepNames = new List<DefinedNameInfo>();
                List<DefinedNameInfo> dropNames = new List<DefinedNameInfo>();
                for (int i = 0; i < definedNames.Count; i++)
                {
                    DefinedNameInfo d = definedNames[i];
                    if (d.Broken) { dropNames.Add(d); report.NameBroken++; continue; }
                    if (usedNames.Contains(d.Name.ToUpperInvariant())) { keepNames.Add(d); report.NameUsed++; continue; }
                    if (d.Name.StartsWith("_xlnm.", StringComparison.Ordinal)) { keepNames.Add(d); report.NameSystem++; continue; }
                    if (opt.Conservative) { keepNames.Add(d); report.NameUnreferenced++; continue; }
                    dropNames.Add(d); report.NameUnreferenced++;
                }
                report.NameTotal = definedNames.Count;

                // ---------------- 解析外部链接 ----------------
                Log(opt, "解析外部链接…");
                List<ExternalRefInfo> extRefs = new List<ExternalRefInfo>();
                foreach (Match m in ExtRefRx.Matches(wb))
                {
                    ExternalRefInfo r = new ExternalRefInfo();
                    r.Raw = m.Value;
                    r.RId = Util.GetAttr(m.Groups[1].Value, "r:id");
                    r.Keep = false;
                    extRefs.Add(r);
                }
                report.LinkTotal = extRefs.Count;

                Dictionary<string, RelationshipInfo> relById = new Dictionary<string, RelationshipInfo>(StringComparer.Ordinal);
                foreach (Match m in RelationshipRx.Matches(relsText))
                {
                    string a = m.Groups[1].Value;
                    RelationshipInfo r = new RelationshipInfo();
                    r.Raw = m.Value;
                    r.Id = Util.GetAttr(a, "Id");
                    r.Type = Util.GetAttr(a, "Type");
                    r.Target = Util.GetAttr(a, "Target");
                    if (!relById.ContainsKey(r.Id)) relById[r.Id] = r;
                }

                // ---------------- 哪些外部链接真的被引用了 ----------------
                HashSet<int> linkByFormula = new HashSet<int>();
                for (int i = 0; i < formulaTexts.Count; i++)
                {
                    string s = WebUtility.HtmlDecode(formulaTexts[i]);
                    s = StringLitRx.Replace(s, "\"\"");   // 字符串常量里的 [n] 不算
                    foreach (Match m in LinkRefRx.Matches(s))
                        linkByFormula.Add(int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture));
                }

                // 保留下来的名称若引用了外部链接，该链接也必须留（否则名称会变成 #REF!）
                HashSet<int> linkByName = new HashSet<int>();
                for (int i = 0; i < keepNames.Count; i++)
                {
                    foreach (Match m in LinkRefRx.Matches(keepNames[i].Value))
                    {
                        int idx = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                        if (!linkByFormula.Contains(idx)) linkByName.Add(idx);
                    }
                }

                for (int i = 0; i < extRefs.Count; i++)
                {
                    if (linkByFormula.Contains(i + 1) || linkByName.Contains(i + 1)) extRefs[i].Keep = true;
                }

                HashSet<string> keptRIds = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < extRefs.Count; i++)
                {
                    if (!extRefs[i].Keep) continue;
                    keptRIds.Add(extRefs[i].RId);
                    if (linkByFormula.Contains(i + 1)) report.LinkKeptByFormula++;
                    else report.LinkKeptByName++;

                    RelationshipInfo rel;
                    if (relById.TryGetValue(extRefs[i].RId, out rel))
                    {
                        string linkPart = rel.Target.StartsWith("/", StringComparison.Ordinal)
                            ? rel.Target.TrimStart('/') : "xl/" + rel.Target;
                        report.KeptLinkTargets.Add(GetExternalLinkRealTarget(map, Util.NormalizePartPath(linkPart)));
                    }
                }
                report.LinkDropped = extRefs.Count - keptRIds.Count;

                // ---------------- 要删除的部件 ----------------
                HashSet<string> dropParts = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < extRefs.Count; i++)
                {
                    if (extRefs[i].Keep) continue;
                    RelationshipInfo rel;
                    if (!relById.TryGetValue(extRefs[i].RId, out rel)) continue;
                    if (rel.Type.IndexOf("externalLink", StringComparison.Ordinal) < 0) continue;
                    string part = rel.Target.StartsWith("/", StringComparison.Ordinal)
                        ? rel.Target.TrimStart('/') : "xl/" + rel.Target;
                    part = Util.NormalizePartPath(part);
                    dropParts.Add(part);
                    dropParts.Add(Regex.Replace(part, "([^/]+)$", "_rels/$1.rels"));
                }

                // 孤儿外部链接部件（workbook.xml 里没有任何 <externalReference> 指向它）
                HashSet<string> keptTargets = new HashSet<string>(StringComparer.Ordinal);
                foreach (string rid in keptRIds)
                {
                    RelationshipInfo rel;
                    if (relById.TryGetValue(rid, out rel))
                        keptTargets.Add(Util.NormalizePartPath("xl/" + rel.Target));
                }
                for (int i = 0; i < srcEntries.Count; i++)
                {
                    string fn = srcEntries[i].FullName;
                    if (!ExtLinkPartRx.IsMatch(fn)) continue;
                    if (keptTargets.Contains(fn)) continue;
                    dropParts.Add(fn);
                    dropParts.Add(Regex.Replace(fn, "([^/]+)$", "_rels/$1.rels"));
                }
                report.DroppedParts = dropParts.Count;

                report.Analyzed = true;
                report.WorkbookXmlAfter = wb.Length;

                if (dropNames.Count == 0 && dropParts.Count == 0)
                {
                    report.NothingToDo = true;
                    report.WorkbookXmlAfter = wb.Length;
                    report.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                    return report;
                }

                if (opt.DryRun)
                {
                    report.WorkbookXmlAfter = wb.Length;
                    report.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                    return report;
                }

                // ---------------- 改写 XML ----------------
                Log(opt, "改写 workbook.xml / rels / Content_Types…");
                HashSet<string> dropNameRaw = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < dropNames.Count; i++) dropNameRaw.Add(dropNames[i].Raw);

                HashSet<string> dropRefRaw = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < extRefs.Count; i++) if (!extRefs[i].Keep) dropRefRaw.Add(extRefs[i].Raw);

                HashSet<string> dropRelRaw = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < extRefs.Count; i++)
                {
                    if (extRefs[i].Keep) continue;
                    RelationshipInfo rel;
                    if (relById.TryGetValue(extRefs[i].RId, out rel)) dropRelRaw.Add(rel.Raw);
                }

                HashSet<string> dropCtRaw = new HashSet<string>(StringComparer.Ordinal);
                foreach (Match m in OverrideRx.Matches(ctText))
                {
                    string pn = Util.GetAttr(m.Groups[1].Value, "PartName").TrimStart('/');
                    if (dropParts.Contains(Util.NormalizePartPath(pn))) dropCtRaw.Add(m.Value);
                }

                string wbNew = RemoveElements(wb, DefinedNameRx, dropNameRaw);
                wbNew = RemoveElements(wbNew, ExtRefRx, dropRefRaw);
                wbNew = EmptyDnRx.Replace(wbNew, "");
                wbNew = EmptyErRx.Replace(wbNew, "");

                if (wbNew.Length >= wb.Length && (dropNames.Count + dropRefRaw.Count) > 0)
                    throw new InvalidOperationException("内部错误：workbook.xml 未被缩短");

                string relsNew = RemoveElements(relsText, RelationshipRx, dropRelRaw);
                string ctNew = RemoveElements(ctText, OverrideRx, dropCtRaw);

                report.WorkbookXmlAfter = wbNew.Length;

                // ---------------- 写包 ----------------
                Log(opt, "写入新包…");
                bool inPlace = opt.InPlace;
                if (inPlace) { tempPath = dst + ".xlsxdoctor.tmp"; workPath = tempPath; }
                else { workPath = dst; }

                if (File.Exists(workPath)) File.Delete(workPath);

                using (FileStream outFs = File.Create(workPath))
                using (ZipArchive outZip = new ZipArchive(outFs, ZipArchiveMode.Create))
                {
                    if (ctPart != null) WriteEntryBytes(outZip, ContentTypesPart, ctPart.ReplaceText(ctNew));

                    for (int i = 0; i < srcEntries.Count; i++)
                    {
                        ZipArchiveEntry e = srcEntries[i];
                        string n = e.FullName;
                        if (n == ContentTypesPart || n == WorkbookPart || n == WorkbookRelsPart) continue;
                        if (dropParts.Contains(n)) continue;
                        if (e.Length == 0 && n.EndsWith("/", StringComparison.Ordinal))
                        {
                            bool stale = false;
                            foreach (string d in dropParts)
                                if (d.StartsWith(n, StringComparison.Ordinal)) { stale = true; break; }
                            if (stale) continue;
                        }
                        CopyEntry(outZip, e);
                    }

                    WriteEntryBytes(outZip, WorkbookPart, wbPart.ReplaceText(wbNew));
                    if (relsPart != null) WriteEntryBytes(outZip, WorkbookRelsPart, relsPart.ReplaceText(relsNew));
                }

                report.WroteFile = true;

                // ---------------- 校验输出 ----------------
                Log(opt, "校验输出包…");
                List<string> errs = PackageValidator.Validate(workPath);
                if (errs.Count > 0)
                {
                    // 很多真实文件本来就不规范（悬空关系、未声明的部件…），而且这些部件本工具根本不碰。
                    // 所以先校验源文件，把源文件自带的同类问题剔除掉：只对「自己引入的」问题负责。
                    Log(opt, "输出包有问题，正在与源文件比对…");
                    List<string> srcErrs = PackageValidator.Validate(fullSrc);
                    HashSet<string> inherited = new HashSet<string>(srcErrs, StringComparer.Ordinal);

                    List<string> introduced = new List<string>();
                    for (int i = 0; i < errs.Count; i++)
                    {
                        if (inherited.Contains(errs[i])) report.PreexistingIssues.Add(errs[i]);
                        else introduced.Add(errs[i]);
                    }

                    if (introduced.Count > 0)
                    {
                        report.ValidationPassed = false;
                        report.ValidationErrors.AddRange(introduced);
                        try { File.Delete(workPath); } catch (Exception) { }
                        report.WroteFile = false;
                        report.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                        return report;
                    }
                }
                report.ValidationPassed = true;

                // ---------------- 落位 ----------------
                if (inPlace)
                {
                    Log(opt, "备份原文件并就地替换…");
                    // 必须先关掉自己打开的源文件流，否则 File.Replace 会因文件被占用而失败
                    try { srcZip.Dispose(); } catch (Exception) { }
                    try { srcFs.Dispose(); } catch (Exception) { }
                    string backup = dst + ".bak";
                    try
                    {
                        if (File.Exists(backup)) File.Delete(backup);
                        File.Replace(workPath, dst, backup);
                        report.BackupPath = backup;
                    }
                    catch (Exception ex)
                    {
                        throw new IOException("无法替换原文件（可能正被 Excel 打开，请先关闭）: " + ex.Message, ex);
                    }
                    workPath = null;
                }

                report.OutputSize = new FileInfo(report.OutputPath).Length;
                report.ElapsedSeconds = sw.Elapsed.TotalSeconds;
                Log(opt, "完成。");
                return report;
            }
            finally
            {
                // 只清理「就地覆盖」用的临时文件；正常模式的输出文件必须留下
                if (tempPath != null && File.Exists(tempPath)) { try { File.Delete(tempPath); } catch (Exception) { } }
                srcZip.Dispose();
                srcFs.Dispose();
            }
        }

        // ============================ 辅助 ============================

        static void Log(CleanOptions opt, string msg)
        {
            if (opt.Log != null) opt.Log(msg);
        }

        /// <summary>单遍正则删除元素。绝不能用逐条 String.Replace：16 万条 × 15MB 文本是 O(n·m)。</summary>
        static string RemoveElements(string text, Regex rx, HashSet<string> dropSet)
        {
            if (dropSet.Count == 0) return text;
            return rx.Replace(text, delegate(Match m)
            {
                return dropSet.Contains(m.Value) ? "" : m.Value;
            });
        }

        static void WriteEntryBytes(ZipArchive zip, string name, byte[] bytes)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream s = e.Open())
                s.Write(bytes, 0, bytes.Length);
        }

        static void CopyEntry(ZipArchive zip, ZipArchiveEntry src)
        {
            ZipArchiveEntry e = zip.CreateEntry(src.FullName, CompressionLevel.Optimal);
            using (Stream outs = e.Open())
            using (Stream ins = src.Open())
            {
                byte[] buf = new byte[65536];
                int n;
                while ((n = ins.Read(buf, 0, buf.Length)) > 0) outs.Write(buf, 0, n);
            }
        }

        /// <summary>
        /// 判断一条公式文本里引用了哪些定义名称。
        ///
        /// 规则遵循「宁可误留、不可误删」：只有能证明「不可能是名称引用」的记号才排除
        /// （字符串常量、单引号工作表名、单元格引用 A1/$B$13、工作表限定符 xxx!、内置函数名）。
        /// </summary>
        static void CollectUsedNames(string formula, HashSet<string> nameLookup,
                                     HashSet<string> lambdaNames, HashSet<string> used)
        {
            string s = WebUtility.HtmlDecode(formula);       // &quot; -> " 等
            s = StringLitRx.Replace(s, "\"\"");              // 字符串常量：里面的名字不算引用
            s = SingleQuotedRx.Replace(s, "''");             // 单引号工作表名：'Sheet 1'!A1

            foreach (Match m in IdentRx.Matches(s))
            {
                string upper = m.Value.ToUpperInvariant();
                int afterIdx = m.Index + m.Length;
                string after = afterIdx >= s.Length ? "" : s.Substring(afterIdx);

                // 工作表限定符 / #REF! 等错误值
                if (after.StartsWith("!", StringComparison.Ordinal)) continue;
                // 单元格引用 A1 / $B$13（Excel 禁止把这种名字定义为名称）
                if (CellTailRx.IsMatch(after)) continue;
                // 更长标识符 / 单元格引用的一部分
                if (m.Index > 0 && IsIdentChar(s[m.Index - 1])) continue;
                // 内置函数名（LAMBDA 定义的名称例外，那是「被调用」）
                if (FuncTailRx.IsMatch(after) && !lambdaNames.Contains(upper)) continue;

                if (nameLookup.Contains(upper)) used.Add(upper);
            }
        }

        /// <summary>
        /// 取外部链接真正指向的文件路径。workbook.xml.rels 里的 Target 只是
        /// "externalLinks/externalLinkN.xml" 这种包内部件名，真正的目标在
        /// xl/externalLinks/_rels/externalLinkN.xml.rels 里（TargetMode="External"）。
        /// </summary>
        static string GetExternalLinkRealTarget(Dictionary<string, ZipArchiveEntry> map, string linkPart)
        {
            string relsName = Regex.Replace(linkPart, "([^/]+)$", "_rels/$1.rels");
            ZipArchiveEntry e;
            if (!map.TryGetValue(relsName, out e)) return linkPart;
            try
            {
                Match m = RelationshipRx.Match(Util.ReadAllText(e));
                if (m.Success)
                {
                    string t = Util.GetAttr(m.Groups[1].Value, "Target");
                    if (t.Length > 0) return t;
                }
            }
            catch (Exception) { }
            return linkPart;
        }

        static bool IsIdentChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '$';
        }
    }

    internal static class TextPartExtensions
    {
        /// <summary>用新文本替换部件内容，保持原编码与 BOM。</summary>
        public static byte[] ReplaceText(this TextPart part, string newText)
        {
            TextPart copy = new TextPart();
            copy.Name = part.Name;
            copy.Encoding = part.Encoding;
            copy.HasBom = part.HasBom;
            copy.Text = newText;
            return copy.Encode();
        }
    }
}
