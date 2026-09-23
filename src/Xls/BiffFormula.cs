using System;
using System.Collections.Generic;
using System.Globalization;

namespace XlsxDoctor.Xls
{
    /// <summary>公式令牌里一处需要重编号的引用。</summary>
    public sealed class TokenRef
    {
        public int Offset;      // 在 rgce 中的字节偏移（令牌首字节）
        public byte Ptg;
        public string Kind = "";
        public int NameIndex = -1;   // ptgName / ptgNameX 的 4 字节索引
        public int Ixti = -1;        // 3D 令牌的 2 字节 ixti（外部表索引）

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "0x{0:X2} {1} @{2} name={3} ixti={4}", Ptg, Kind, Offset, NameIndex, Ixti);
        }
    }

    /// <summary>
    /// BIFF8 公式令牌流（rgce）的解析与改写。
    ///
    /// 这是整个 .xls 清理里最容易出错的一环：公式里的定义名称和外部引用
    /// 不是按名字存的，而是按**索引号**存的。
    ///   删掉 NAME 记录 → 所有 ptgName 的索引都要跟着挪，否则公式全指错。
    ///   删掉 SUPBOOK/EXTERNSHEET → 所有 ptgRef3d/ptgArea3d 的 ixti 也要挪。
    ///
    /// 令牌首字节：低 5 位是编号，bit5-6 是"类"（0=无类 / 1=值 / 2=引用 / 3=数组）。
    /// 类不改变长度，所以长度表按 ptg & 0x1F 查即可；但 <0x20 的无类令牌要单独处理。
    /// </summary>
    public static class BiffFormula
    {
        // ---- 类无关令牌（ptg < 0x20）----
        // 大部分是运算符/括号，长度 1；少数带操作数。
        static readonly int[] UnclassedFixed = BuildUnclassedFixed();

        static int[] BuildUnclassedFixed()
        {
            int[] t = new int[0x20];
            for (int i = 0; i < t.Length; i++) t[i] = -1;

            t[0x01] = 5;   // ptgExp      : ptg + rw(2) + col(2)
            t[0x02] = 5;   // ptgTbl
            for (int i = 0x03; i <= 0x11; i++) t[i] = 1;   // 运算符 / 区间 / 并集 / 交集
            t[0x12] = 1;   // ptgUplus
            t[0x13] = 1;   // ptgUminus
            t[0x14] = 1;   // ptgPercent
            t[0x15] = 1;   // ptgParen
            t[0x16] = 1;   // ptgMissArg
            t[0x17] = -2;  // ptgStr      : 变长
            t[0x19] = -3;  // ptgAttr     : 通常是 4，fChoose 时变长
            t[0x1C] = 2;   // ptgErr
            t[0x1D] = 2;   // ptgBool
            t[0x1E] = 3;   // ptgInt      : ptg + u16
            t[0x1F] = 9;   // ptgNum      : ptg + f64
            return t;
        }

        /// <summary>
        /// 类令牌（ptg >= 0x20）的长度，按 ptg & 0x1F 查。
        /// 注意 0x19/0x1A/0x1B/0x1C/0x1D 在类区间里分别是 NameX / Ref3d / Area3d / RefErr3d / AreaErr3d，
        /// 与无类区间的 ptgAttr 等编号重合，靠"类位"区分。
        /// </summary>
        static int ClassedLength(int id)
        {
            switch (id)
            {
                case 0x00: return 8;    // ptgArray     : ptg + 7 字节保留
                case 0x01: return 3;    // ptgFunc      : ptg + iftab(2)
                case 0x02: return 4;    // ptgFuncVar   : ptg + argc(1) + iftab(2)
                case 0x03: return 5;    // ptgName      : ptg + nameindex(4)
                case 0x04: return 5;    // ptgRef       : ptg + rw(2) + col(2)
                case 0x05: return 9;    // ptgArea
                case 0x06: return 7;    // ptgMemArea   : ptg + reserved(4) + cce(2)
                case 0x07: return 7;    // ptgMemErr
                case 0x08: return 7;    // ptgMemNoMem
                case 0x09: return 5;    // ptgMemFunc
                case 0x0A: return 5;    // ptgRefErr
                case 0x0B: return 9;    // ptgAreaErr
                case 0x0C: return 5;    // ptgRefN
                case 0x0D: return 9;    // ptgAreaN
                case 0x0E: return 7;    // ptgMemAreaN
                case 0x0F: return 7;    // ptgMemNoMemN
                case 0x19: return 7;    // ptgNameX     : ptg + ixti(2) + nameindex(4)
                case 0x1A: return 7;    // ptgRef3d     : ptg + ixti(2) + rw(2) + col(2)
                case 0x1B: return 11;   // ptgArea3d    : ptg + ixti(2) + rwFirst..colLast(8)
                case 0x1C: return 7;    // ptgRefErr3d
                case 0x1D: return 11;   // ptgAreaErr3d
                default: return -1;
            }
        }

        /// <summary>单条令牌的字节长度。返回 -1 表示无法识别的令牌。</summary>
        public static int LengthOfToken(byte[] rgce, int pos, out string name)
        {
            byte ptg = rgce[pos];

            if (ptg < 0x20)
            {
                int fixedLen = UnclassedFixed[ptg];
                if (fixedLen > 0)
                {
                    name = UnclassedName(ptg);
                    return fixedLen;
                }
                if (fixedLen == -2)
                {
                    name = "ptgStr";
                    // BIFF8：ptg(1) + cch(1) + grbit(1) + 字符
                    if (pos + 3 > rgce.Length) return -1;
                    int cch = rgce[pos + 1];
                    int grbit = rgce[pos + 2];
                    int bytes = ((grbit & 0x01) != 0) ? cch * 2 : cch;
                    return 3 + bytes;
                }
                if (fixedLen == -3)
                {
                    name = "ptgAttr";
                    // ptg(1) + grbit(1) + data(2)；fChoose(bit2) 时后面跟一张跳转表
                    if (pos + 4 > rgce.Length) return -1;
                    int grbit = rgce[pos + 1];
                    if ((grbit & 0x04) != 0)
                    {
                        // ptg + grbit + cce(1) + rgOffset(2*(cce+1))
                        if (pos + 3 > rgce.Length) return -1;
                        int cnt = rgce[pos + 2];
                        return 1 + 1 + 1 + 2 * (cnt + 1);
                    }
                    return 4;
                }
                name = "ptg?0x" + ptg.ToString("X2", CultureInfo.InvariantCulture);
                return -1;
            }

            int id = ptg & 0x1F;
            int len = ClassedLength(id);
            if (len < 0)
            {
                name = "ptg?0x" + ptg.ToString("X2", CultureInfo.InvariantCulture);
                return -1;
            }
            name = ClassedName(id);
            return len;
        }

        static string UnclassedName(byte ptg)
        {
            switch (ptg)
            {
                case 0x01: return "ptgExp";
                case 0x02: return "ptgTbl";
                case 0x03: return "ptgAdd";
                case 0x04: return "ptgSub";
                case 0x05: return "ptgMul";
                case 0x06: return "ptgDiv";
                case 0x07: return "ptgPower";
                case 0x08: return "ptgConcat";
                case 0x09: return "ptgLT";
                case 0x0A: return "ptgLE";
                case 0x0B: return "ptgEQ";
                case 0x0C: return "ptgGE";
                case 0x0D: return "ptgGT";
                case 0x0E: return "ptgNE";
                case 0x0F: return "ptgIsect";
                case 0x10: return "ptgUnion";
                case 0x11: return "ptgRange";
                case 0x12: return "ptgUplus";
                case 0x13: return "ptgUminus";
                case 0x14: return "ptgPercent";
                case 0x15: return "ptgParen";
                case 0x16: return "ptgMissArg";
                case 0x1C: return "ptgErr";
                case 0x1D: return "ptgBool";
                case 0x1E: return "ptgInt";
                case 0x1F: return "ptgNum";
                default: return "ptg?";
            }
        }

        static string ClassedName(int id)
        {
            switch (id)
            {
                case 0x00: return "ptgArray";
                case 0x01: return "ptgFunc";
                case 0x02: return "ptgFuncVar";
                case 0x03: return "ptgName";
                case 0x04: return "ptgRef";
                case 0x05: return "ptgArea";
                case 0x06: return "ptgMemArea";
                case 0x07: return "ptgMemErr";
                case 0x08: return "ptgMemNoMem";
                case 0x09: return "ptgMemFunc";
                case 0x0A: return "ptgRefErr";
                case 0x0B: return "ptgAreaErr";
                case 0x0C: return "ptgRefN";
                case 0x0D: return "ptgAreaN";
                case 0x0E: return "ptgMemAreaN";
                case 0x0F: return "ptgMemNoMemN";
                case 0x19: return "ptgNameX";
                case 0x1A: return "ptgRef3d";
                case 0x1B: return "ptgArea3d";
                case 0x1C: return "ptgRefErr3d";
                case 0x1D: return "ptgAreaErr3d";
                default: return "ptg?";
            }
        }

        // ------------------------------------------------------------------ 扫描

        /// <summary>
        /// 扫一遍 rgce，返回所有需要重编号的引用。
        /// 解析失败（遇到不认识的令牌或越界）返回 false 并给出原因 —— 调用方必须据此放弃清理，
        /// 绝不能"猜"着往下走，否则会把公式改坏。
        /// </summary>
        public static bool TryScan(byte[] rgce, int cce, List<TokenRef> refs, out string error)
        {
            refs.Clear();
            error = null;

            int pos = 0;
            int guard = 0;
            while (pos < cce)
            {
                if (guard++ > 200000) { error = "令牌数量异常（疑似死循环）"; return false; }

                string tname;
                int len = LengthOfToken(rgce, pos, out tname);
                if (len < 0) { error = string.Format(CultureInfo.InvariantCulture, "偏移 {0} 处遇到无法识别的令牌 0x{1:X2}", pos, rgce[pos]); return false; }
                if (pos + len > cce) { error = string.Format(CultureInfo.InvariantCulture, "偏移 {0} 处令牌 {1} 长度 {2} 越出 rgce（cce={3}）", pos, tname, len, cce); return false; }

                byte ptg = rgce[pos];
                TokenRef tr = null;

                if (ptg >= 0x20)
                {
                    int id = ptg & 0x1F;
                    if (id == 0x03)
                    {
                        tr = new TokenRef();
                        tr.Kind = "ptgName";
                        tr.NameIndex = (int)BitConverter.ToUInt32(rgce, pos + 1);
                    }
                    else if (id == 0x19)
                    {
                        tr = new TokenRef();
                        tr.Kind = "ptgNameX";
                        tr.Ixti = BitConverter.ToUInt16(rgce, pos + 1);
                        tr.NameIndex = (int)BitConverter.ToUInt32(rgce, pos + 3);
                    }
                    else if (id == 0x1A || id == 0x1B || id == 0x1C || id == 0x1D)
                    {
                        tr = new TokenRef();
                        tr.Kind = ClassedName(id);
                        tr.Ixti = BitConverter.ToUInt16(rgce, pos + 1);
                    }
                }

                if (tr != null)
                {
                    tr.Offset = pos;
                    tr.Ptg = ptg;
                    refs.Add(tr);
                }

                pos += len;
            }

            if (pos != cce)
            {
                error = string.Format(CultureInfo.InvariantCulture, "令牌流长度不符：走完 {0} 字节，cce={1}", pos, cce);
                return false;
            }
            return true;
        }

        /// <summary>按映射表原地改写 rgce 里的索引。令牌长度不变，所以可以直接覆盖写。</summary>
        public static void Remap(byte[] rgce, List<TokenRef> refs, int[] nameMap, int[] ixtiMap)
        {
            for (int i = 0; i < refs.Count; i++)
            {
                TokenRef tr = refs[i];

                if (tr.NameIndex >= 0 && nameMap != null)
                {
                    int nv = (tr.NameIndex < nameMap.Length) ? nameMap[tr.NameIndex] : -1;
                    if (nv < 0)
                        throw new XlsFormatException(
                            "内部错误：公式引用了本应被删除的名称索引 " + tr.NameIndex.ToString(CultureInfo.InvariantCulture) + "，这是不允许发生的");
                    rgce[tr.Offset + 1] = (byte)(nv & 0xFF);
                    rgce[tr.Offset + 2] = (byte)((nv >> 8) & 0xFF);
                    rgce[tr.Offset + 3] = (byte)((nv >> 16) & 0xFF);
                    rgce[tr.Offset + 4] = (byte)((nv >> 24) & 0xFF);
                }

                if (tr.Ixti >= 0 && ixtiMap != null)
                {
                    int iv = (tr.Ixti < ixtiMap.Length) ? ixtiMap[tr.Ixti] : -1;
                    if (iv < 0)
                        throw new XlsFormatException(
                            "内部错误：公式引用了本应被删除的外部表索引 " + tr.Ixti.ToString(CultureInfo.InvariantCulture) + "，这是不允许发生的");
                    rgce[tr.Offset + 1] = (byte)(iv & 0xFF);
                    rgce[tr.Offset + 2] = (byte)((iv >> 8) & 0xFF);
                }
            }
        }

        // ------------------------------------------------------------------ FORMULA 记录布局

        /// <summary>
        /// BIFF8 FORMULA(0x0006) 记录头长度：
        /// rw(2) col(2) ixfe(2) val(8) grbit(2) chn(4) cce(2) = 22 字节，rgce 从 22 开始。
        /// </summary>
        public const int FormulaHeaderLength = 22;

        /// <summary>从 FORMULA 记录数据里取出 rgce 偏移与长度。失败返回 false。</summary>
        public static bool TryGetRgce(byte[] data, out int rgceOffset, out int cce, out string error)
        {
            rgceOffset = 0;
            cce = 0;
            error = null;

            if (data.Length < FormulaHeaderLength)
            {
                error = "FORMULA 记录太短（" + data.Length.ToString(CultureInfo.InvariantCulture) + " 字节）";
                return false;
            }

            cce = BitConverter.ToUInt16(data, 20);
            rgceOffset = FormulaHeaderLength;
            if (rgceOffset + cce > data.Length)
            {
                error = string.Format(CultureInfo.InvariantCulture,
                    "cce={0} 超出记录长度 {1}", cce, data.Length);
                return false;
            }
            return true;
        }
    }
}
