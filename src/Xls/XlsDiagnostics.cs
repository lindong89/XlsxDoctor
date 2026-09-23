using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace XlsxDoctor.Xls
{
    /// <summary>
    /// .xls 结构诊断（只读，不修改文件）。
    /// 输出格式刻意与开发期用的 PowerShell 原型 analyze-xls.ps1 对齐，便于交叉验证。
    /// </summary>
    public static class XlsDiagnostics
    {
        public static string Describe(string path)
        {
            return Describe(path, 14);
        }

        public static string Describe(string path, int topN)
        {
            FileInfo fi = new FileInfo(path);
            StringBuilder sb = new StringBuilder();

            CfbfFile cfbf = CfbfFile.Read(path);

            sb.AppendLine("文件          : " + path);
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "大小          : {0:N0} 字节", fi.Length));
            sb.AppendLine("CFBF 版本     : " + cfbf.MajorVersion.ToString(CultureInfo.InvariantCulture)
                          + "（扇区 " + cfbf.SectorSize.ToString(CultureInfo.InvariantCulture)
                          + " 字节 / mini " + cfbf.MiniSectorSize.ToString(CultureInfo.InvariantCulture) + " 字节）");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "FAT 扇区数    : {0}", cfbf.NumFatSectors));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "mini 阈值     : {0} 字节", cfbf.MiniStreamCutoff));
            sb.AppendLine();

            sb.AppendLine("---- 目录项 ----");
            for (int i = 0; i < cfbf.Entries.Count; i++)
            {
                CfbfEntry e = cfbf.Entries[i];
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "  [{0,-7}] #{1,-3} {2,-30} {3,14:N0} 字节",
                    e.TypeName, e.Index, e.Name, e.Size));
            }
            sb.AppendLine();

            CfbfEntry wbEntry = cfbf.FindWorkbookEntry();
            if (wbEntry == null)
            {
                sb.AppendLine("目录里找不到 Workbook / Book 流 —— 这不像是 Excel 工作簿。");
                return sb.ToString();
            }

            byte[] stream = cfbf.ReadStream(wbEntry);
            BiffFile biff = BiffFile.Parse(stream);

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "Workbook 流   : {0:N0} 字节", stream.Length));

            int bofCount = 0;
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].Type == 0x0809) bofCount++;

            // CONTINUE 续块在 BiffFile 里已并入父记录，所以这里的条数是"逻辑记录"数。
            // 同时报出物理条数，便于跟按物理记录统计的工具对拍。
            int contTotal = 0;
            for (int i = 0; i < biff.Records.Count; i++) contTotal += biff.Records[i].ContinuationCount;

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "BIFF 记录     : {0:N0} 条逻辑记录 = {1:N0} 条物理记录（其中 {2:N0} 条是 CONTINUE 续块）",
                biff.Records.Count, biff.Records.Count + contTotal, contTotal));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "                （BOF 子流 {0} 个 = 工作表/图表/VBA 等）", bofCount));

            // 最大使用范围
            int maxRow = 0, maxCol = 0, dims = 0;
            for (int i = 0; i < biff.Records.Count; i++)
            {
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x0200 || r.Data.Length < 12) continue;
                dims++;
                uint rwMac = BitConverter.ToUInt32(r.Data, 4);
                ushort colMac = BitConverter.ToUInt16(r.Data, 10);
                if (rwMac > maxRow) maxRow = (int)rwMac;
                if (colMac > maxCol) maxCol = colMac;
            }
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "最大使用范围  : {0} 行 x {1} 列   （DIMENSIONS 记录 {2} 条）", maxRow, maxCol, dims));
            if (maxRow >= 65536) sb.AppendLine("  !! 触及 .xls 的 65536 行上限 —— 转档时可能已经丢数据");
            if (maxCol >= 256) sb.AppendLine("  !! 触及 .xls 的 256 列上限 —— 转档时可能已经丢数据");
            sb.AppendLine();

            Dictionary<int, int> counts = new Dictionary<int, int>();
            Dictionary<int, long> bytes = new Dictionary<int, long>();
            biff.Tally(counts, bytes);

            sb.AppendLine("---- 按数量排序（前 " + topN.ToString(CultureInfo.InvariantCulture) + "）----");
            Dictionary<int, long> countsAsLong = new Dictionary<int, long>();
            foreach (KeyValuePair<int, int> kv in counts) countsAsLong[kv.Key] = kv.Value;
            AppendRanked(sb, countsAsLong, stream.Length, topN, false);
            sb.AppendLine();

            sb.AppendLine("---- 按占用字节排序（前 " + topN.ToString(CultureInfo.InvariantCulture) + "）----");
            AppendRanked(sb, bytes, stream.Length, topN, true);
            sb.AppendLine();

            sb.AppendLine("---- 最大的记录块（父记录 + 跟随的 CONTINUE）----");
            List<BiffRecord> runs = new List<BiffRecord>();
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].ContinuationCount > 0) runs.Add(biff.Records[i]);
            runs.Sort(delegate (BiffRecord a, BiffRecord b) { return b.TotalLength.CompareTo(a.TotalLength); });
            for (int i = 0; i < runs.Count && i < 8; i++)
            {
                BiffRecord r = runs[i];
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "  0x{0:X4}  {1,-14} {2,12:N0} 字节  {3,5:N1}%  (含 {4} 个 CONTINUE)",
                    r.Type, r.TypeName, r.TotalLength,
                    100.0 * r.TotalLength / stream.Length, r.ContinuationCount));
            }
            sb.AppendLine();

            // MSODRAWINGGROUP（内嵌图片）
            BiffRecord dg = null;
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].Type == 0x00EB) { dg = biff.Records[i]; break; }

            if (dg != null)
            {
                sb.AppendLine("---- MSODRAWINGGROUP 内部（Escher，" +
                              dg.Data.Length.ToString("N0", CultureInfo.InvariantCulture) + " 字节）----");
                DescribeEscher(sb, dg.Data);
                sb.AppendLine();
            }

            // ---- 定义名称样例 ----
            List<NameRecordInfo> nm = XlsCleaner.ParseNames(biff);
            sb.AppendLine();
            sb.AppendLine("---- 定义名称（前 " + topN.ToString(CultureInfo.InvariantCulture) + " 条，共 "
                          + nm.Count.ToString("N0", CultureInfo.InvariantCulture) + " 条）----");
            for (int i = 0; i < nm.Count && i < topN; i++)
            {
                NameRecordInfo n = nm[i];
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    #{0,-6} grbit=0x{1:X4} chKey=0x{2:X2} itab={3,-4} 名称=[{4}]",
                    n.Index0, n.Grbit, n.ChKey, n.Itab, n.Name));
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "            cce={0,-5} 名称字节={1,-4} 系统名={2,-6} 错误令牌={3,-6} 解析={4}",
                    n.Cce, n.NameBytes, n.IsSystemName, n.HasErrorToken, n.Parsed));
                sb.AppendLine("            rgch 原始: " + n.RawNameHex);
            }

            sb.AppendLine();
            sb.AppendLine("---- 垃圾相关记录 ----");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  NAME        (0x0018 定义名称)  : {0,10:N0} 条，本身 {1:N0} 字节，含续块 {2:N0} 字节",
                biff.CountOf(0x0018), biff.BytesOfSelf(0x0018), biff.BytesOf(0x0018)));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  EXTERNSHEET (0x0017 外部表)    : {0,10:N0} 条，本身 {1:N0} 字节，含续块 {2:N0} 字节",
                biff.CountOf(0x0017), biff.BytesOfSelf(0x0017), biff.BytesOf(0x0017)));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  EXTERNNAME  (0x0023 外部名称)  : {0,10:N0} 条，本身 {1:N0} 字节，含续块 {2:N0} 字节",
                biff.CountOf(0x0023), biff.BytesOfSelf(0x0023), biff.BytesOf(0x0023)));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  SUPBOOK     (0x01AE 外部工作簿): {0,10:N0} 条，本身 {1:N0} 字节，含续块 {2:N0} 字节",
                biff.CountOf(0x01AE), biff.BytesOfSelf(0x01AE), biff.BytesOf(0x01AE)));
            sb.AppendLine();

            UsageScan usage = BiffUsage.Scan(biff);
            sb.Append(BiffUsage.Describe(usage));
            sb.AppendLine();

            RefTable rt = RefTable.Parse(biff);
            rt.MarkReferenced(usage.ReferencedIxti);
            sb.Append(rt.Describe(12));

            return sb.ToString();
        }

        static void AppendRanked(StringBuilder sb, Dictionary<int, long> map, int streamLen, int topN, bool isBytes)
        {
            List<KeyValuePair<int, long>> list = new List<KeyValuePair<int, long>>(map);
            list.Sort(delegate (KeyValuePair<int, long> a, KeyValuePair<int, long> b)
            {
                return b.Value.CompareTo(a.Value);
            });
            for (int i = 0; i < list.Count && i < topN; i++)
            {
                int t = list[i].Key;
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "  0x{0:X4}  {1,-14} {2,12:N0}{3}",
                    t, BiffRecordNames.Get((ushort)t), list[i].Value,
                    isBytes ? string.Format(CultureInfo.InvariantCulture, "  {0,5:N1}%", 100.0 * list[i].Value / streamLen) : ""));
            }
        }

        /// <summary>遍历 Escher 记录，统计类型，并读出每张内嵌图片的原始大小。</summary>
        static void DescribeEscher(StringBuilder sb, byte[] dg)
        {
            Dictionary<int, int> escCounts = new Dictionary<int, int>();
            List<long> bseSizes = new List<long>();

            int q = 0;
            int guard = 0;
            while (q + 8 <= dg.Length && guard < 1000000)
            {
                ushort verInst = BitConverter.ToUInt16(dg, q);
                ushort etype = BitConverter.ToUInt16(dg, q + 2);
                uint elen = BitConverter.ToUInt32(dg, q + 4);
                int ver = verInst & 0x000F;
                bool isContainer = (ver == 0x000F);

                int c;
                escCounts.TryGetValue(etype, out c);
                escCounts[etype] = c + 1;

                // BSE 布局：btWin32(1) btMacOS(1) rgbUid(16) tag(2) size(4) ...
                if (etype == 0xF007 && elen >= 36)
                    bseSizes.Add(BitConverter.ToUInt32(dg, q + 8 + 20));

                if (isContainer) q += 8;
                else q += 8 + (int)elen;
                guard++;
            }

            List<KeyValuePair<int, int>> list = new List<KeyValuePair<int, int>>(escCounts);
            list.Sort(delegate (KeyValuePair<int, int> a, KeyValuePair<int, int> b)
            {
                return b.Value.CompareTo(a.Value);
            });

            sb.AppendLine("  Escher 记录类型（前 10）：");
            for (int i = 0; i < list.Count && i < 10; i++)
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    0x{0:X4}  {1,-18} {2,8:N0}", list[i].Key, EscherName(list[i].Key), list[i].Value));

            long sum = 0;
            for (int i = 0; i < bseSizes.Count; i++) sum += bseSizes[i];
            sb.AppendLine();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  内嵌图片（BSE）数量 : {0} 个", bseSizes.Count));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  图片原始数据合计    : {0:N0} 字节  ({1:N2} MB)", sum, sum / 1048576.0));
            if (bseSizes.Count > 0)
            {
                bseSizes.Sort();
                bseSizes.Reverse();
                sb.AppendLine("  单张图片大小（前 10，降序）：");
                for (int i = 0; i < bseSizes.Count && i < 10; i++)
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "    {0,12:N0} 字节", bseSizes[i]));
            }
        }

        static string EscherName(int t)
        {
            switch (t)
            {
                case 0xF000: return "DggContainer";
                case 0xF001: return "BstoreContainer";
                case 0xF002: return "DgContainer";
                case 0xF003: return "SpgrContainer";
                case 0xF004: return "SpContainer";
                case 0xF006: return "Dgg";
                case 0xF007: return "BSE(图片)";
                case 0xF008: return "FDGG";
                case 0xF009: return "FBSE";
                case 0xF00A: return "FDG";
                case 0xF00B: return "FSPGR";
                case 0xF00C: return "FSP";
                case 0xF00D: return "FOPT";
                case 0xF00F: return "SecondaryFOPT";
                case 0xF010: return "TertiaryFOPT";
                case 0xF118: return "FOPT3";
                case 0xF11A: return "ClientAnchor";
                case 0xF11E: return "ClientTextbox";
                case 0xF122: return "ClientData";
                default: return "?";
            }
        }
    }
}
