using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace XlsxDoctor.Xls
{
    /// <summary>.xls 文件结构异常。</summary>
    public sealed class XlsFormatException : Exception
    {
        public XlsFormatException(string message) : base(message) { }
    }

    /// <summary>CFBF 目录项（128 字节）。</summary>
    public sealed class CfbfEntry
    {
        public int Index;
        public string Name = "";
        public byte ObjectType;     // 0=未用 1=Storage 2=Stream 5=Root
        public byte ColorFlag;
        public int LeftSibling;
        public int RightSibling;
        public int Child;
        public uint StateBits;
        public uint StartSector;
        public long Size;

        /// <summary>原始 128 字节。重建容器时原样写回，这样目录树的红黑结构不会被动坏。</summary>
        public byte[] Raw = new byte[128];

        public bool IsStream { get { return ObjectType == 2; } }
        public bool IsRoot { get { return ObjectType == 5; } }

        public string TypeName
        {
            get
            {
                if (ObjectType == 1) return "Storage";
                if (ObjectType == 2) return "Stream";
                if (ObjectType == 5) return "Root";
                return "Other";
            }
        }
    }

    /// <summary>
    /// CFBF（Compound File Binary Format，即 OLE2 复合文档）容器。
    ///
    /// .xls 就是这个容器，里面装着一个名为 "Workbook" 的流（BIFF 记录序列）。
    ///
    /// 布局：512 字节头部 → 扇区数组。
    ///   FAT        描述每个扇区的后继（扇区链）
    ///   目录项     每个 128 字节，描述一个流/存储的名字、起始扇区、长度
    ///   mini 流    小于 4096 字节的流不占整扇区，统一塞进 Root Entry 的流里，
    ///              由 mini FAT 描述
    ///   DIFAT      当 FAT 本身超过 109 个扇区时，用 DIFAT 链继续指向 FAT 扇区
    ///
    /// 注意：特殊扇区号必须用 uint 比较。写成十六进制字面量再当 int 比较是错的
    /// （0xFFFFFFFE 会被当成 -2），这个坑在 PowerShell 原型里踩过一次。
    /// </summary>
    public sealed class CfbfFile
    {
        public const uint FREESECT   = 4294967295u;   // 0xFFFFFFFF
        public const uint ENDOFCHAIN = 4294967294u;   // 0xFFFFFFFE
        public const uint FATSECT    = 4294967293u;   // 0xFFFFFFFD
        public const uint DIFSECT    = 4294967292u;   // 0xFFFFFFFC

        private static readonly byte[] Signature =
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public byte[] Raw;
        public ushort MajorVersion;
        public ushort ByteOrder;
        public int SectorSize;
        public int MiniSectorSize;
        public uint FirstDirSector;
        public uint MiniStreamCutoff;
        public uint FirstMiniFatSector;
        public uint NumMiniFatSectors;
        public uint FirstDifatSector;
        public uint NumDifatSectors;
        public uint NumFatSectors;

        public List<uint> Fat = new List<uint>();
        public List<uint> MiniFat = new List<uint>();
        public List<CfbfEntry> Entries = new List<CfbfEntry>();
        public byte[] MiniStream = new byte[0];

        /// <summary>扇区 n 在文件中的字节偏移（头部占前 512 字节，相当于"第 -1 扇区"）。</summary>
        public long SectorOffset(uint n)
        {
            return ((long)n + 1) * SectorSize;
        }

        // ------------------------------------------------------------------ 读取

        public static CfbfFile Read(string path)
        {
            return Parse(File.ReadAllBytes(path));
        }

        public static CfbfFile Parse(byte[] raw)
        {
            if (raw.Length < 512)
                throw new XlsFormatException("文件太小（不足 512 字节），不可能是 .xls");

            for (int i = 0; i < 8; i++)
            {
                if (raw[i] != Signature[i])
                    throw new XlsFormatException(
                        "不是 CFBF/OLE2 复合文档（文件签名不匹配）。" +
                        "如果这其实是 .xlsx，请把扩展名改回去。");
            }

            CfbfFile f = new CfbfFile();
            f.Raw = raw;
            f.MajorVersion = BitConverter.ToUInt16(raw, 0x1A);
            f.ByteOrder = BitConverter.ToUInt16(raw, 0x1C);

            int sectorShift = BitConverter.ToUInt16(raw, 0x1E);
            int miniShift = BitConverter.ToUInt16(raw, 0x20);

            if (sectorShift < 7 || sectorShift > 20)
                throw new XlsFormatException("CFBF 扇区大小字段异常（sectorShift=" + sectorShift + "）");
            if (miniShift < 2 || miniShift > 20)
                throw new XlsFormatException("CFBF mini 扇区大小字段异常（miniShift=" + miniShift + "）");

            f.SectorSize = 1 << sectorShift;
            f.MiniSectorSize = 1 << miniShift;
            f.NumFatSectors = BitConverter.ToUInt32(raw, 0x2C);
            f.FirstDirSector = BitConverter.ToUInt32(raw, 0x30);
            f.MiniStreamCutoff = BitConverter.ToUInt32(raw, 0x38);
            f.FirstMiniFatSector = BitConverter.ToUInt32(raw, 0x3C);
            f.NumMiniFatSectors = BitConverter.ToUInt32(raw, 0x40);
            f.FirstDifatSector = BitConverter.ToUInt32(raw, 0x44);
            f.NumDifatSectors = BitConverter.ToUInt32(raw, 0x48);

            if (f.MiniStreamCutoff == 0) f.MiniStreamCutoff = 4096;

            f.ReadFat();
            f.ReadDirectory();
            f.ReadMiniFat();
            f.ReadMiniStream();
            return f;
        }

        /// <summary>收集 FAT 扇区号（先读头部内嵌的 109 项，再沿 DIFAT 链扩展），然后读出整张 FAT。</summary>
        private void ReadFat()
        {
            List<uint> fatSectors = new List<uint>();
            for (int i = 0; i < 109; i++)
            {
                uint s = BitConverter.ToUInt32(Raw, 0x4C + i * 4);
                if (s == FREESECT) break;
                fatSectors.Add(s);
            }

            uint next = FirstDifatSector;
            int guard = 0;
            while (next != ENDOFCHAIN && next != FREESECT && guard < 100000)
            {
                long off = SectorOffset(next);
                if (off + SectorSize > Raw.Length) break;
                int cnt = (SectorSize / 4) - 1;
                for (int i = 0; i < cnt; i++)
                {
                    uint s = BitConverter.ToUInt32(Raw, (int)(off + i * 4));
                    if (s == FREESECT) break;
                    fatSectors.Add(s);
                }
                next = BitConverter.ToUInt32(Raw, (int)(off + cnt * 4));
                guard++;
            }

            int perSector = SectorSize / 4;
            for (int i = 0; i < fatSectors.Count; i++)
            {
                long off = SectorOffset(fatSectors[i]);
                if (off + SectorSize > Raw.Length) continue;
                for (int k = 0; k < perSector; k++)
                    Fat.Add(BitConverter.ToUInt32(Raw, (int)(off + k * 4)));
            }
        }

        private void ReadDirectory()
        {
            byte[] dir = ReadChain(FirstDirSector, 0);
            for (int i = 0; i + 128 <= dir.Length; i += 128)
            {
                byte objType = dir[i + 0x42];
                if (objType == 0) continue;

                CfbfEntry e = new CfbfEntry();
                e.Index = i / 128;
                int nameLen = BitConverter.ToUInt16(dir, i + 0x40);
                if (nameLen > 2 && nameLen <= 64)
                    e.Name = Encoding.Unicode.GetString(dir, i, nameLen - 2);
                e.ObjectType = objType;
                e.ColorFlag = dir[i + 0x43];
                e.LeftSibling = (int)BitConverter.ToUInt32(dir, i + 0x44);
                e.RightSibling = (int)BitConverter.ToUInt32(dir, i + 0x48);
                e.Child = (int)BitConverter.ToUInt32(dir, i + 0x4C);
                e.StateBits = BitConverter.ToUInt32(dir, i + 0x60);
                e.StartSector = BitConverter.ToUInt32(dir, i + 0x74);

                uint sizeLo = BitConverter.ToUInt32(dir, i + 0x78);
                if (MajorVersion >= 4)
                {
                    uint sizeHi = BitConverter.ToUInt32(dir, i + 0x7C);
                    e.Size = (long)sizeLo | ((long)sizeHi << 32);
                }
                else
                {
                    e.Size = sizeLo;
                }

                Buffer.BlockCopy(dir, i, e.Raw, 0, 128);
                Entries.Add(e);
            }
        }

        private void ReadMiniFat()
        {
            if (NumMiniFatSectors == 0 || FirstMiniFatSector == ENDOFCHAIN) return;
            byte[] mf = ReadChain(FirstMiniFatSector, 0);
            for (int i = 0; i + 4 <= mf.Length; i += 4)
                MiniFat.Add(BitConverter.ToUInt32(mf, i));
        }

        private void ReadMiniStream()
        {
            CfbfEntry root = FindRoot();
            if (root == null || root.Size == 0) { MiniStream = new byte[0]; return; }
            MiniStream = ReadChain(root.StartSector, root.Size);
        }

        /// <summary>沿 FAT 链读出一整段流。size 为 0 表示读到链尾（用于 FAT/目录这类自身长度不定的结构）。</summary>
        public byte[] ReadChain(uint start, long size)
        {
            MemoryStream ms = new MemoryStream();
            uint s = start;
            int guard = 0;
            while (s != ENDOFCHAIN && s != FREESECT && guard < 4000000)
            {
                if (s >= (uint)Fat.Count) break;
                long off = SectorOffset(s);
                if (off >= Raw.Length) break;
                long len = Math.Min((long)SectorSize, (long)Raw.Length - off);
                if (len <= 0) break;
                ms.Write(Raw, (int)off, (int)len);
                s = Fat[(int)s];
                guard++;
            }
            return Truncate(ms, size);
        }

        /// <summary>沿 mini FAT 链读出一小段流（从 MiniStream 里切）。</summary>
        public byte[] ReadMiniChain(uint start, long size)
        {
            MemoryStream ms = new MemoryStream();
            uint s = start;
            int guard = 0;
            while (s != ENDOFCHAIN && s != FREESECT && guard < 4000000)
            {
                if (s >= (uint)MiniFat.Count) break;
                long off = (long)s * MiniSectorSize;
                if (off >= MiniStream.Length) break;
                long len = Math.Min((long)MiniSectorSize, (long)MiniStream.Length - off);
                if (len <= 0) break;
                ms.Write(MiniStream, (int)off, (int)len);
                s = MiniFat[(int)s];
                guard++;
            }
            return Truncate(ms, size);
        }

        private static byte[] Truncate(MemoryStream ms, long size)
        {
            byte[] all = ms.ToArray();
            ms.Dispose();
            if (size <= 0 || all.Length <= size) return all;
            byte[] cut = new byte[size];
            Buffer.BlockCopy(all, 0, cut, 0, (int)size);
            return cut;
        }

        /// <summary>读出一个目录项对应的流数据（自动区分普通扇区链与 mini 流）。</summary>
        public byte[] ReadStream(CfbfEntry e)
        {
            if (e == null || e.Size == 0) return new byte[0];
            if (e.Size < MiniStreamCutoff) return ReadMiniChain(e.StartSector, e.Size);
            return ReadChain(e.StartSector, e.Size);
        }

        public CfbfEntry FindEntry(string name)
        {
            for (int i = 0; i < Entries.Count; i++)
            {
                if (string.Equals(Entries[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return Entries[i];
            }
            return null;
        }

        public CfbfEntry FindRoot()
        {
            for (int i = 0; i < Entries.Count; i++)
                if (Entries[i].IsRoot) return Entries[i];
            return null;
        }

        /// <summary>取 BIFF 工作簿流。BIFF8 叫 "Workbook"，BIFF5 叫 "Book"。</summary>
        public CfbfEntry FindWorkbookEntry()
        {
            CfbfEntry e = FindEntry("Workbook");
            if (e == null) e = FindEntry("Book");
            return e;
        }

        // ------------------------------------------------------------------ 诊断输出

        public string Describe()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("CFBF 版本     : " + MajorVersion.ToString(CultureInfo.InvariantCulture)
                          + "（扇区 " + SectorSize.ToString(CultureInfo.InvariantCulture)
                          + " 字节 / mini " + MiniSectorSize.ToString(CultureInfo.InvariantCulture) + " 字节）");
            sb.AppendLine("FAT 扇区数    : " + NumFatSectors.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("mini 阈值     : " + MiniStreamCutoff.ToString(CultureInfo.InvariantCulture) + " 字节");
            sb.AppendLine("目录项        : " + Entries.Count.ToString(CultureInfo.InvariantCulture) + " 个");
            sb.AppendLine();
            sb.AppendLine("---- 目录项 ----");
            for (int i = 0; i < Entries.Count; i++)
            {
                CfbfEntry e = Entries[i];
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "  [{0,-7}] #{1,-3} {2,-30} {3,14:N0} 字节",
                    e.TypeName, e.Index, e.Name, e.Size));
            }
            return sb.ToString();
        }
    }
}
