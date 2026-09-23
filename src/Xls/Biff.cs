using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace XlsxDoctor.Xls
{
    /// <summary>
    /// 一条 BIFF 记录（含它后面紧跟的所有 CONTINUE 续块）。
    ///
    /// BIFF 记录格式：[类型 u16][长度 u16][数据 长度 字节]。
    /// 单条记录数据上限 8224 字节，超出部分拆成若干 CONTINUE(0x003C) 记录紧跟其后。
    /// 这里把续块合并回一条逻辑记录，Data 是拼接后的完整数据；
    /// 重建时按 Start/TotalLength 原样搬运原始字节，避免重新分段的麻烦。
    /// </summary>
    public sealed class BiffRecord
    {
        public ushort Type;
        public int Start;             // 记录头在流中的偏移
        public int TotalLength;       // 含所有 CONTINUE 头的总字节数
        public int ContinuationCount;
        /// <summary>记录本身（第一段）的数据长度，不含续块。用于区分'记录本身占用'与'含续块占用'。</summary>
        public int FirstSegmentLength;
        public byte[] Data;           // 合并续块后的数据

        /// <summary>
        /// 改写后的数据。非 null 表示这条记录要重新打包写出。
        /// 必须 <= 8224 字节（本工具只改写 FORMULA / BOUNDSHEET 这类小记录，
        /// 不需要重新做 CONTINUE 分段）。
        /// </summary>
        public byte[] NewData;

        public bool IsModified { get { return NewData != null; } }

        public string TypeName { get { return BiffRecordNames.Get(Type); } }

        public string Describe()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "0x{0:X4} {1,-14} {2,10:N0} 字节 (含 {3} 个 CONTINUE)",
                Type, TypeName, TotalLength, ContinuationCount);
        }
    }

    /// <summary>BIFF 记录类型名。</summary>
    public static class BiffRecordNames
    {
        static readonly Dictionary<int, string> Map = BuildMap();

        static Dictionary<int, string> BuildMap()
        {
            Dictionary<int, string> m = new Dictionary<int, string>();
            m[0x0006] = "FORMULA";      m[0x000A] = "EOF";          m[0x0017] = "EXTERNSHEET";
            m[0x0018] = "NAME";         m[0x0023] = "EXTERNNAME";   m[0x003C] = "CONTINUE";
            m[0x0085] = "BOUNDSHEET";   m[0x00FC] = "SST";          m[0x00FD] = "LABELSST";
            m[0x0203] = "NUMBER";       m[0x027E] = "RK";           m[0x0204] = "LABEL";
            m[0x01AE] = "SUPBOOK";      m[0x0809] = "BOF";          m[0x00E0] = "XF";
            m[0x0031] = "FONT";         m[0x041E] = "FORMAT";       m[0x00E5] = "MERGEDCELLS";
            m[0x00BD] = "MULRK";        m[0x00FF] = "EXTSST";       m[0x0092] = "PALETTE";
            m[0x023E] = "WINDOW2";      m[0x0055] = "DEFCOLWIDTH";  m[0x00D7] = "DBCELL";
            m[0x0208] = "ROW";          m[0x0201] = "BLANK";        m[0x00BE] = "MULBLANK";
            m[0x00EB] = "MSODRAWGRP";   m[0x00EC] = "MSODRAWING";   m[0x005D] = "OBJ";
            m[0x01B6] = "TXO";          m[0x087C] = "XFCRC";        m[0x087D] = "XFEXT";
            m[0x007D] = "COLINFO";      m[0x0200] = "DIMENSIONS";   m[0x00E1] = "INTERFACEHDR";
            m[0x008C] = "COUNTRY";      m[0x0022] = "DATEMODE";     m[0x0042] = "CODEPAGE";
            m[0x013D] = "RRTABID";      m[0x009C] = "BUILTINFMTCOUNT"; m[0x00A0] = "FONTX";
            m[0x0422] = "HLINK";        m[0x005C] = "WRITEACCESS";  m[0x013C] = "CONTINUEFRT";
            m[0x0862] = "SHEETLAYOUT";  m[0x0863] = "SHEETPROTECTION";
            m[0x005E] = "UNCALCED";     m[0x0099] = "STANDARDWIDTH"; m[0x0225] = "DEFAULTROWHEIGHT";
            m[0x0059] = "XCT";          m[0x005A] = "CRN";
            m[0x00DA] = "BOOKBOOL";     m[0x00A1] = "SETUP";        m[0x00A9] = "GRIDSET";
            m[0x0207] = "STRING";       m[0x0004] = "LABEL(BIFF2)";
            m[0x020B] = "INDEX";        m[0x0231] = "FILESHARING";
            return m;
        }

        public static string Get(ushort type)
        {
            string s;
            if (Map.TryGetValue(type, out s)) return s;
            return "?";
        }
    }

    /// <summary>BIFF8 工作簿流（Workbook 流）的记录序列。</summary>
    public sealed class BiffFile
    {
        public const int MaxRecordData = 8224;

        public byte[] Stream;
        public List<BiffRecord> Records = new List<BiffRecord>();

        public static BiffFile Parse(byte[] stream)
        {
            BiffFile f = new BiffFile();
            f.Stream = stream;

            int p = 0;
            while (p + 4 <= stream.Length)
            {
                ushort rt = BitConverter.ToUInt16(stream, p);
                ushort rl = BitConverter.ToUInt16(stream, p + 2);
                if (p + 4 + rl > stream.Length) break;   // 尾部残缺，就此打住

                int start = p;
                p += 4 + rl;

                // 合并紧跟其后的 CONTINUE 续块
                MemoryStream ms = new MemoryStream();
                ms.Write(stream, start + 4, rl);
                int contCount = 0;
                while (p + 4 <= stream.Length)
                {
                    ushort nt = BitConverter.ToUInt16(stream, p);
                    if (nt != 0x003C) break;
                    ushort nl = BitConverter.ToUInt16(stream, p + 2);
                    if (p + 4 + nl > stream.Length) break;
                    ms.Write(stream, p + 4, nl);
                    contCount++;
                    p += 4 + nl;
                }

                BiffRecord r = new BiffRecord();
                r.Type = rt;
                r.Start = start;
                r.TotalLength = p - start;
                r.ContinuationCount = contCount;
                r.FirstSegmentLength = rl;
                r.Data = ms.ToArray();
                ms.Dispose();
                f.Records.Add(r);
            }

            return f;
        }

        /// <summary>统计各记录类型出现的次数与占用字节数。</summary>
        public void Tally(Dictionary<int, int> counts, Dictionary<int, long> bytes)
        {
            counts.Clear();
            bytes.Clear();
            for (int i = 0; i < Records.Count; i++)
            {
                int t = Records[i].Type;
                int c;
                counts.TryGetValue(t, out c);
                counts[t] = c + 1;

                long b;
                bytes.TryGetValue(t, out b);
                bytes[t] = b + Records[i].TotalLength;
            }
        }

        public int CountOf(ushort type)
        {
            int n = 0;
            for (int i = 0; i < Records.Count; i++)
                if (Records[i].Type == type) n++;
            return n;
        }

        public long BytesOf(ushort type)
        {
            long n = 0;
            for (int i = 0; i < Records.Count; i++)
                if (Records[i].Type == type) n += Records[i].TotalLength;
            return n;
        }

        /// <summary>只算记录本身（不含 CONTINUE 续块）占用的字节数。</summary>
        public long BytesOfSelf(ushort type)
        {
            long n = 0;
            for (int i = 0; i < Records.Count; i++)
                if (Records[i].Type == type) n += 4 + Records[i].FirstSegmentLength;
            return n;
        }

        /// <summary>
        /// 按 drop 集合重建整条工作簿流。未被删除的记录原样搬运原始字节，
        /// 被改写的记录（NewData 非 null）重新打包成单条记录。
        /// </summary>
        public byte[] Rebuild(HashSet<int> dropIndices)
        {
            Dictionary<int, int> dummy;
            return Rebuild(dropIndices, out dummy);
        }

        /// <summary>
        /// 重建整条流，同时给出"原偏移 -> 新偏移"的映射。
        /// BOUNDSHEET 的 lbPlyPos 和 EXTSST 的分桶偏移都是流内绝对偏移，
        /// 删掉字节之后必须靠这张表重新算，否则工作表定位就错位了。
        /// </summary>
        public byte[] Rebuild(HashSet<int> dropIndices, out Dictionary<int, int> offsetMap)
        {
            offsetMap = new Dictionary<int, int>();
            MemoryStream outMs = new MemoryStream(Stream.Length);
            for (int i = 0; i < Records.Count; i++)
            {
                if (dropIndices != null && dropIndices.Contains(i)) continue;

                BiffRecord r = Records[i];
                offsetMap[r.Start] = (int)outMs.Length;
                if (r.NewData == null)
                {
                    outMs.Write(Stream, r.Start, r.TotalLength);
                }
                else
                {
                    if (r.NewData.Length > MaxRecordData)
                        throw new XlsFormatException(
                            "内部错误：改写后的记录超过 8224 字节，需要重新分段（本工具不应出现这种情况）");
                    byte[] hdr = new byte[4];
                    hdr[0] = (byte)(r.Type & 0xFF);
                    hdr[1] = (byte)((r.Type >> 8) & 0xFF);
                    hdr[2] = (byte)(r.NewData.Length & 0xFF);
                    hdr[3] = (byte)((r.NewData.Length >> 8) & 0xFF);
                    outMs.Write(hdr, 0, 4);
                    outMs.Write(r.NewData, 0, r.NewData.Length);
                }
            }
            byte[] all = outMs.ToArray();
            outMs.Dispose();
            return all;
        }
    }
}
