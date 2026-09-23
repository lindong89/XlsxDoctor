using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace XlsxDoctor.Xls
{
    /// <summary>一条 FORMULA 记录的解析结果。</summary>
    public sealed class FormulaInfo
    {
        public int RecordIndex;
        public int RgceOffset;
        public int Cce;
        public List<TokenRef> Refs = new List<TokenRef>();
    }

    /// <summary>整条工作簿流的"使用情况"扫描结果。</summary>
    public sealed class UsageScan
    {
        public int FormulaRecords;
        public int ParseOk;
        public int ParseFail;
        public List<string> FailSamples = new List<string>();

        public List<FormulaInfo> Formulas = new List<FormulaInfo>();

        public HashSet<int> ReferencedNameIndices = new HashSet<int>();
        public HashSet<int> ReferencedIxti = new HashSet<int>();

        public int PtgNameRefs;
        public int PtgNameXRefs;
        public int ThreeDRefs;

        /// <summary>出现过的令牌首字节 → 次数。用来验证长度表是否覆盖了实际用到的全部令牌。</summary>
        public Dictionary<int, int> TokenHistogram = new Dictionary<int, int>();

        public int NameRecordCount;

        /// <summary>裸字节交叉校验：在整条流里按 ptgName 字节模式扫到的名称索引。</summary>
        public HashSet<int> RawPtgNameHits = new HashSet<int>();

        /// <summary>裸扫命中但正规解析没发现的索引个数。&gt;0 表示还有没考虑到的引用位置，必须保守处理。</summary>
        public int RawHitsNotParsed;

        /// <summary>裸扫命中（且正规解析未发现）按所在记录类型统计 —— 用来判断这些命中是真引用还是图像数据里的噪声。</summary>
        public Dictionary<ushort, int> RawHitsByRecordType = new Dictionary<ushort, int>();

        public List<string> RawHitSamples = new List<string>();

        /// <summary>FORMULA 记录里的裸命中详情：判断命中落在 rgce 内（真令牌）还是 rgce 之后的数据区（噪声）。</summary>
        public List<string> FormulaHitDetail = new List<string>();
        public int FormulaHitsInsideRgce;
        public int FormulaHitsOutsideRgce;

        /// <summary>落在 rgce 内的命中：把该记录的 rgce 完整展开，逐令牌核对。</summary>
        public List<string> RgceInsideHitDump = new List<string>();

        public bool AllParsed { get { return ParseFail == 0; } }
    }

    /// <summary>
    /// 扫描整条工作簿流，弄清"哪些定义名称 / 哪些外部表真的被公式引用"。
    ///
    /// 安全原则：只要有**任何一条**公式解析不出来，就整体放弃清理。
    /// 宁可不动，也不能在解析不确定的情况下改公式索引。
    /// </summary>
    public static class BiffUsage
    {
        public static UsageScan Scan(BiffFile biff)
        {
            UsageScan s = new UsageScan();

            for (int i = 0; i < biff.Records.Count; i++)
            {
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x0006) continue;

                s.FormulaRecords++;

                int rgceOffset, cce;
                string err;
                if (!BiffFormula.TryGetRgce(r.Data, out rgceOffset, out cce, out err))
                {
                    s.ParseFail++;
                    if (s.FailSamples.Count < 5)
                        s.FailSamples.Add("记录 #" + i.ToString(CultureInfo.InvariantCulture) + "（记录 " + r.Start.ToString(CultureInfo.InvariantCulture) + " 字节处）：" + err);
                    continue;
                }

                byte[] rgce = Slice(r.Data, rgceOffset, cce);

                List<TokenRef> refs = new List<TokenRef>();
                if (!BiffFormula.TryScan(rgce, cce, refs, out err))
                {
                    s.ParseFail++;
                    if (s.FailSamples.Count < 5)
                        s.FailSamples.Add("记录 #" + i.ToString(CultureInfo.InvariantCulture) + "（流偏移 " + r.Start.ToString(CultureInfo.InvariantCulture) + "）：" + err);
                    continue;
                }

                s.ParseOk++;

                FormulaInfo fi = new FormulaInfo();
                fi.RecordIndex = i;
                fi.RgceOffset = rgceOffset;
                fi.Cce = cce;
                fi.Refs = refs;
                s.Formulas.Add(fi);

                // 令牌直方图（顺带验证长度表覆盖了实际出现的全部令牌）
                int pos = 0;
                while (pos < cce)
                {
                    int b = rgce[pos];
                    int c;
                    s.TokenHistogram.TryGetValue(b, out c);
                    s.TokenHistogram[b] = c + 1;

                    string tname;
                    int len = BiffFormula.LengthOfToken(rgce, pos, out tname);
                    if (len <= 0) break;
                    pos += len;
                }

                for (int k = 0; k < refs.Count; k++)
                {
                    TokenRef tr = refs[k];
                    if (tr.Kind == "ptgName") { s.PtgNameRefs++; s.ReferencedNameIndices.Add(tr.NameIndex); }
                    else if (tr.Kind == "ptgNameX") { s.PtgNameXRefs++; s.ReferencedNameIndices.Add(tr.NameIndex); s.ReferencedIxti.Add(tr.Ixti); }
                    else { s.ThreeDRefs++; s.ReferencedIxti.Add(tr.Ixti); }
                }
            }

            // ---- 裸字节交叉校验 ----
            // 正规解析只覆盖 FORMULA 记录。为了确认没有"藏在别的记录里"的 ptgName
            // （比如 SHRFMLA 共享公式、图表子流里的系列公式），把整条流按字节模式再扫一遍：
            // ptgName 的编码是 [0x23|0x43|0x63][4 字节小端索引]，索引必须落在 [0, 名称数) 内。
            // 只要有解析结果之外的命中，就说明存在盲区，清理逻辑必须退让。
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].Type == 0x0018) s.NameRecordCount++;

            if (s.NameRecordCount > 0 && biff.Stream != null)
            {
                byte[] st = biff.Stream;
                for (int ri = 0; ri < biff.Records.Count; ri++)
                {
                    BiffRecord rec = biff.Records[ri];
                    int end = rec.Start + rec.TotalLength;
                    if (end > st.Length) end = st.Length;
                    for (int i = rec.Start; i + 5 <= end; i++)
                    {
                        byte b = st[i];
                        if (b != 0x23 && b != 0x43 && b != 0x63) continue;
                        uint v = BitConverter.ToUInt32(st, i + 1);
                        if (v >= (uint)s.NameRecordCount) continue;

                        int vi = (int)v;
                        s.RawPtgNameHits.Add(vi);

                        if (!s.ReferencedNameIndices.Contains(vi))
                        {
                            int c;
                            s.RawHitsByRecordType.TryGetValue(rec.Type, out c);
                            s.RawHitsByRecordType[rec.Type] = c + 1;
                            if (s.RawHitSamples.Count < 12)
                                s.RawHitSamples.Add(string.Format(CultureInfo.InvariantCulture,
                                    "索引 #{0,-8} 落在 0x{1:X4} {2,-14}（记录内偏移 {3}）",
                                    vi, rec.Type, rec.TypeName, i - rec.Start));

                            if (rec.Type == 0x0006)
                            {
                                int rgceOff2, cce2; string e2;
                                if (BiffFormula.TryGetRgce(rec.Data, out rgceOff2, out cce2, out e2))
                                {
                                    int dataOff = (i - rec.Start) - 4;
                                    bool inside = (dataOff >= rgceOff2) && (dataOff + 5 <= rgceOff2 + cce2);
                                    if (inside) s.FormulaHitsInsideRgce++; else s.FormulaHitsOutsideRgce++;

                                    if (inside && s.RgceInsideHitDump.Count < 3)
                                    {
                                        byte[] rg = new byte[cce2];
                                        Buffer.BlockCopy(rec.Data, rgceOff2, rg, 0, cce2);
                                        StringBuilder d = new StringBuilder();
                                        d.AppendLine("记录 #" + ri.ToString(CultureInfo.InvariantCulture)
                                                     + "  rgce 共 " + cce2.ToString(CultureInfo.InvariantCulture) + " 字节");
                                        d.Append("    原始字节: ");
                                        for (int k = 0; k < rg.Length; k++) d.Append(rg[k].ToString("X2", CultureInfo.InvariantCulture) + " ");
                                        d.AppendLine();
                                        d.AppendLine("    逐令牌解析：");
                                        int p3 = 0;
                                        while (p3 < rg.Length)
                                        {
                                            string tn3; int tl3 = BiffFormula.LengthOfToken(rg, p3, out tn3);
                                            d.AppendLine("      偏移 " + p3.ToString(CultureInfo.InvariantCulture)
                                                         + ": 0x" + rg[p3].ToString("X2", CultureInfo.InvariantCulture)
                                                         + " " + tn3 + " 长度 " + tl3.ToString(CultureInfo.InvariantCulture));
                                            if (tl3 <= 0) break;
                                            p3 += tl3;
                                        }
                                        int hitRel = dataOff - rgceOff2;
                                        d.AppendLine("    命中在 rgce 内偏移 " + hitRel.ToString(CultureInfo.InvariantCulture)
                                                     + "，该字节 = 0x" + rg[hitRel].ToString("X2", CultureInfo.InvariantCulture));
                                        s.RgceInsideHitDump.Add(d.ToString());
                                    }
                                    if (s.FormulaHitDetail.Count < 12)
                                        s.FormulaHitDetail.Add(string.Format(CultureInfo.InvariantCulture,
                                            "记录 #{0} 数据长 {1} cce={2} 命中在数据偏移 {3} → {4}",
                                            ri, rec.Data.Length, cce2, dataOff, inside ? "★在 rgce 内" : "在 rgce 之后（rgcb 数据区）"));
                                }
                            }
                        }
                    }
                }
                foreach (int v in s.RawPtgNameHits)
                    if (!s.ReferencedNameIndices.Contains(v)) s.RawHitsNotParsed++;
            }

            return s;
        }

        /// <summary>把 TryScan 的长度表开放出来给直方图用。</summary>
        static byte[] Slice(byte[] src, int offset, int len)
        {
            byte[] d = new byte[len];
            Buffer.BlockCopy(src, offset, d, 0, len);
            return d;
        }

        public static string Describe(UsageScan s)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("---- 公式令牌解析自检 ----");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  FORMULA 记录      : {0:N0} 条", s.FormulaRecords));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  解析成功          : {0:N0} 条 ({1:N1}%)", s.ParseOk,
                s.FormulaRecords == 0 ? 100.0 : 100.0 * s.ParseOk / s.FormulaRecords));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  解析失败          : {0:N0} 条", s.ParseFail));

            if (s.FailSamples.Count > 0)
            {
                sb.AppendLine("  失败样例：");
                for (int i = 0; i < s.FailSamples.Count; i++)
                    sb.AppendLine("    " + s.FailSamples[i]);
            }

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  ptgName  引用     : {0:N0} 处，涉及 {1:N0} 个名称索引",
                s.PtgNameRefs, s.ReferencedNameIndices.Count));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  ptgNameX 引用     : {0:N0} 处", s.PtgNameXRefs));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  3D 引用(ref/area) : {0:N0} 处，涉及 {1:N0} 个外部表索引",
                s.ThreeDRefs, s.ReferencedIxti.Count));

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  名称记录总数      : {0:N0} 条", s.NameRecordCount));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  裸字节扫到的名称索引: {0:N0} 个，其中正规解析未发现的 {1:N0} 个",
                s.RawPtgNameHits.Count, s.RawHitsNotParsed));
            if (s.RawHitsNotParsed > 0)
            {
                sb.AppendLine("    未发现命中所在的记录类型（按命中数降序）：");
                List<KeyValuePair<ushort, int>> byType = new List<KeyValuePair<ushort, int>>(s.RawHitsByRecordType);
                byType.Sort(delegate (KeyValuePair<ushort, int> a, KeyValuePair<ushort, int> b) { return b.Value.CompareTo(a.Value); });
                for (int i = 0; i < byType.Count && i < 10; i++)
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "      0x{0:X4} {1,-14} {2,8:N0} 处", byType[i].Key,
                        BiffRecordNames.Get(byType[i].Key), byType[i].Value));
                sb.AppendLine("    样例：");
                for (int i = 0; i < s.RawHitSamples.Count; i++)
                    sb.AppendLine("      " + s.RawHitSamples[i]);

                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    FORMULA 内命中：rgce 内 {0:N0} 处，rgce 之后 {1:N0} 处",
                    s.FormulaHitsInsideRgce, s.FormulaHitsOutsideRgce));
                for (int i = 0; i < s.FormulaHitDetail.Count; i++)
                    sb.AppendLine("      " + s.FormulaHitDetail[i]);

                for (int i = 0; i < s.RgceInsideHitDump.Count; i++)
                    sb.Append(s.RgceInsideHitDump[i]);
            }
            sb.AppendLine();
            sb.AppendLine("  令牌直方图（前 20）：");
            List<KeyValuePair<int, int>> hist = new List<KeyValuePair<int, int>>(s.TokenHistogram);
            hist.Sort(delegate (KeyValuePair<int, int> a, KeyValuePair<int, int> b) { return b.Value.CompareTo(a.Value); });
            for (int i = 0; i < hist.Count && i < 20; i++)
            {
                string nm;
                byte[] one = new byte[] { (byte)hist[i].Key, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                BiffFormula.LengthOfToken(one, 0, out nm);
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    0x{0:X2} {1,-14} {2,10:N0}", hist[i].Key, nm, hist[i].Value));
            }

            return sb.ToString();
        }
    }
}
