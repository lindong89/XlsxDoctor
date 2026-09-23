using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace XlsxDoctor.Xls
{
    /// <summary>一条 SUPBOOK 记录（外部工作簿声明，或本工作簿自引用）。</summary>
    public sealed class SupbookInfo
    {
        public int RecordIndex;    // 在 BiffFile.Records 里的下标
        public int Index;          // 在 SUPBOOK 序列中的序号 = 公式/EXTERNSHEET 里的 iSupBook
        public ushort Cch;         // 特殊标记：0x0401=自引用 0x3A01=加载项 0x0000=未用
        public int Ctab;           // 工作表数（自引用/加载项时有效）
        public string Path = "";   // 外部工作簿路径
        public int ExtraBytes;     // 记录里除头部外的剩余字节
        public string RawHead = "";   // 记录前 24 字节的十六进制，用于核对头部布局
        public int DataLength;

        public bool IsSelfReferencing { get { return Cch == 0x0401; } }
        public bool IsAddIn { get { return Cch == 0x3A01; } }
        public bool IsUnused { get { return Cch == 0x0000; } }

        public string KindName
        {
            get
            {
                if (IsSelfReferencing) return "自引用(本工作簿)";
                if (IsAddIn) return "加载项";
                if (IsUnused) return "未使用";
                return "外部工作簿";
            }
        }
    }

    /// <summary>EXTERNSHEET 里的一条 (iSupBook, itabFirst, itabLast) 三元组。</summary>
    public sealed class ExternSheetEntry
    {
        public int Ixti;           // 就是公式 ptgRef3d/ptgArea3d 里的 ixti
        public ushort ISupBook;
        public ushort ItabFirst;
        public ushort ItabLast;
        public bool ReferencedByFormula;

        public string Describe()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "ixti={0,-6} iSupBook={1,-6} itabFirst={2,-6} itabLast={3,-6} {4}",
                Ixti, ISupBook, ItabFirst, ItabLast, ReferencedByFormula ? "★被公式引用" : "");
        }
    }

    /// <summary>
    /// 名称表 + 外部引用表的完整模型。
    ///
    /// 这是安全清理的核心：BIFF 里公式按**索引**引用名称和外部表，
    /// 所以删除任何一条都必须同步重编号。这个类负责把"谁引用了谁"理清楚。
    /// </summary>
    public sealed class RefTable
    {
        public List<SupbookInfo> Supbooks = new List<SupbookInfo>();
        public List<ExternSheetEntry> ExternSheet = new List<ExternSheetEntry>();

        /// <summary>自引用（虚拟）SUPBOOK 的下标；-1 表示没有。它代表本工作簿自己，不是外部链接。</summary>
        public int VirtualSupbookIndex = -1;

        public static RefTable Parse(BiffFile biff)
        {
            RefTable t = new RefTable();

            // ---- SUPBOOK ----
            for (int i = 0; i < biff.Records.Count; i++)
            {
                BiffRecord r = biff.Records[i];
                if (r.Type != 0x01AE) continue;

                SupbookInfo sb = new SupbookInfo();
                sb.RecordIndex = i;
                sb.Index = t.Supbooks.Count;
                sb.DataLength = r.Data.Length;

                int hx = Math.Min(24, r.Data.Length);
                for (int q = 0; q < hx; q++)
                    sb.RawHead += r.Data[q].ToString("X2", CultureInfo.InvariantCulture) + " ";

                if (r.Data.Length >= 2)
                {
                    sb.Cch = BitConverter.ToUInt16(r.Data, 0);

                    // 虚拟 SUPBOOK（自引用 / 加载项）的标记 0x0401 / 0x3A01 可能落在偏移 0，也可能落在偏移 2：
                    // 不同写出器（Excel / WPS / Go Excelize）对 ctab 与 cch 的先后顺序不一致。
                    // 两种布局都必须认 —— 漏判会把本工作簿的三维引用误当成"外部链接"删掉，
                    // 那是最危险的一类误删（公式会静默指向别的表）。
                    ushort v0 = BitConverter.ToUInt16(r.Data, 0);
                    ushort v2 = (r.Data.Length >= 4) ? BitConverter.ToUInt16(r.Data, 2) : (ushort)0;

                    if (v0 == 0x0401 || v0 == 0x3A01)
                    {
                        sb.Cch = v0;
                        sb.Ctab = v2;
                        sb.ExtraBytes = r.Data.Length - 4;
                    }
                    else if (v2 == 0x0401 || v2 == 0x3A01)
                    {
                        sb.Cch = v2;
                        sb.Ctab = v0;
                        sb.ExtraBytes = r.Data.Length - 4;
                    }
                    else if (v0 == 0x0000)
                    {
                        sb.Cch = 0;
                        sb.ExtraBytes = r.Data.Length - 2;
                    }
                    else
                    {
                        // grbit(1) + rgch(cch 字节或 2*cch 字节)
                        if (r.Data.Length >= 3)
                        {
                            int grbit = r.Data[2];
                            int cch = sb.Cch;
                            int need = ((grbit & 0x01) != 0) ? cch * 2 : cch;
                            int avail = Math.Min(need, r.Data.Length - 3);
                            if (avail > 0)
                            {
                                if ((grbit & 0x01) != 0)
                                    sb.Path = Encoding.Unicode.GetString(r.Data, 3, avail);
                                else
                                {
                                    char[] cs = new char[avail];
                                    for (int k = 0; k < avail; k++) cs[k] = (char)r.Data[3 + k];
                                    sb.Path = new string(cs);
                                }
                            }
                        }
                        sb.ExtraBytes = r.Data.Length - 3 - (((r.Data.Length >= 3) && ((r.Data[2] & 0x01) != 0)) ? sb.Cch * 2 : sb.Cch);
                    }
                }

                t.Supbooks.Add(sb);
            }

            // 记下自引用（虚拟）SUPBOOK 的下标 —— 它代表"本工作簿自己"，
            // 是三维引用 Sheet1:Sheet3!A1 的基础，属于正常内容而非垃圾。
            for (int i = 0; i < t.Supbooks.Count; i++)
            {
                if (t.Supbooks[i].IsSelfReferencing) { t.VirtualSupbookIndex = i; break; }
            }

            // ---- EXTERNSHEET ----
            BiffRecord es = null;
            for (int i = 0; i < biff.Records.Count; i++)
                if (biff.Records[i].Type == 0x0017) { es = biff.Records[i]; break; }

            if (es != null && es.Data.Length >= 2)
            {
                int cXti = BitConverter.ToUInt16(es.Data, 0);
                int max = (es.Data.Length - 2) / 6;
                if (cXti > max) cXti = max;   // 容错：声明数超过实际字节数时按实际算

                for (int k = 0; k < cXti; k++)
                {
                    ExternSheetEntry e = new ExternSheetEntry();
                    e.Ixti = k;
                    e.ISupBook = BitConverter.ToUInt16(es.Data, 2 + k * 6);
                    e.ItabFirst = BitConverter.ToUInt16(es.Data, 2 + k * 6 + 2);
                    e.ItabLast = BitConverter.ToUInt16(es.Data, 2 + k * 6 + 4);
                    t.ExternSheet.Add(e);
                }
            }

            return t;
        }

        /// <summary>按公式实际引用情况标注 EXTERNSHEET 条目。</summary>
        public void MarkReferenced(HashSet<int> referencedIxti)
        {
            for (int i = 0; i < ExternSheet.Count; i++)
                ExternSheet[i].ReferencedByFormula = referencedIxti.Contains(ExternSheet[i].Ixti);
        }

        public string Describe(int sampleLimit)
        {
            StringBuilder sb = new StringBuilder();

            Dictionary<string, int> kinds = new Dictionary<string, int>();
            for (int i = 0; i < Supbooks.Count; i++)
            {
                string k = Supbooks[i].KindName;
                int c;
                kinds.TryGetValue(k, out c);
                kinds[k] = c + 1;
            }

            sb.AppendLine("---- SUPBOOK 构成 ----");
            foreach (KeyValuePair<string, int> kv in kinds)
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "  {0,-18} {1,8:N0} 条", kv.Key, kv.Value));

            sb.AppendLine("  cch 字段分布（前 8）：");
            Dictionary<int, int> cchHist = new Dictionary<int, int>();
            for (int i = 0; i < Supbooks.Count; i++)
            {
                int c = Supbooks[i].Cch;
                int n2;
                cchHist.TryGetValue(c, out n2);
                cchHist[c] = n2 + 1;
            }
            List<KeyValuePair<int, int>> chl = new List<KeyValuePair<int, int>>(cchHist);
            chl.Sort(delegate (KeyValuePair<int, int> a, KeyValuePair<int, int> b) { return b.Value.CompareTo(a.Value); });
            for (int i = 0; i < chl.Count && i < 8; i++)
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    cch=0x{0:X4} ({1}) : {2:N0} 条", chl[i].Key, chl[i].Key, chl[i].Value));

            sb.AppendLine("  样例（前 " + sampleLimit.ToString(CultureInfo.InvariantCulture) + " 条）：");
            for (int i = 0; i < Supbooks.Count && i < sampleLimit; i++)
            {
                SupbookInfo s = Supbooks[i];
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "    #{0,-5} {1,-18} cch=0x{2:X4} ctab={3,-5} 长度={4,-7} 头部={5}",
                    s.Index, s.KindName, s.Cch, s.Ctab, s.DataLength, s.RawHead));
                if (s.Path.Length > 0)
                    sb.AppendLine("           路径=" + (s.Path.Length > 60 ? s.Path.Substring(0, 60) + "…" : s.Path));
            }

            sb.AppendLine();
            sb.AppendLine("---- EXTERNSHEET 条目 ----");
            int refd = 0;
            for (int i = 0; i < ExternSheet.Count; i++)
                if (ExternSheet[i].ReferencedByFormula) refd++;

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  共 {0:N0} 条，其中被公式引用 {1:N0} 条", ExternSheet.Count, refd));

            int toVirtual = 0;
            for (int i = 0; i < ExternSheet.Count; i++)
                if (VirtualSupbookIndex >= 0 && ExternSheet[i].ISupBook == VirtualSupbookIndex) toVirtual++;
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "  自引用 SUPBOOK 下标 : {0}（指向本工作簿自己），引用它的条目 {1:N0} 条",
                VirtualSupbookIndex, toVirtual));

            sb.AppendLine("  被公式引用的条目：");
            int shown = 0;
            for (int i = 0; i < ExternSheet.Count && shown < 30; i++)
            {
                if (!ExternSheet[i].ReferencedByFormula) continue;
                ExternSheetEntry e = ExternSheet[i];
                string sbk = (e.ISupBook < Supbooks.Count) ? Supbooks[e.ISupBook].KindName : "越界";
                sb.AppendLine("    " + e.Describe() + "  ->  " + sbk);
                shown++;
            }
            if (refd > shown)
                sb.AppendLine("    …（其余 " + (refd - shown).ToString(CultureInfo.InvariantCulture) + " 条略）");

            return sb.ToString();
        }
    }
}
