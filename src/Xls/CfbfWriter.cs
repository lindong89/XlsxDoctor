using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace XlsxDoctor.Xls
{
    /// <summary>
    /// 重建 CFBF 容器。
    ///
    /// 策略：**保留原目录项的 128 字节原样**，只改写其中的"起始扇区"和"流长度"两个字段。
    /// 这样目录里的红黑树指针（左/右兄弟、子节点下标）完全不动，不用担心树结构被改坏 ——
    /// 那些指针指的是"目录项下标"，不是扇区号，所以重排扇区不影响它们。
    ///
    /// 扇区布局（顺序固定，便于校验）：
    ///   0 .. 大流         Workbook 等 ≥4096 字节的流
    ///   接着  mini 流容器  Root Entry 的流，把所有 <4096 字节的小流串起来
    ///   接着  mini FAT
    ///   接着  目录
    ///   接着  FAT
    ///   接着  DIFAT（FAT 超过 109 个扇区时才需要）
    /// </summary>
    public static class CfbfWriter
    {
        /// <summary>
        /// 写出一个新的 .xls。replacements 以目录项下标为键，值为替换后的流数据；
        /// 未列出的流原样保留。
        /// </summary>
        public static void Write(string path, CfbfFile src, Dictionary<int, byte[]> replacements)
        {
            int sectorSize = src.SectorSize;
            int miniSize = src.MiniSectorSize;
            uint cutoff = src.MiniStreamCutoff;

            // ---------- 1. 取出每个流的最终数据 ----------
            byte[][] data = new byte[src.Entries.Count][];
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                if (!e.IsStream && !e.IsRoot) { data[i] = null; continue; }

                byte[] rep;
                if (replacements != null && replacements.TryGetValue(i, out rep))
                    data[i] = rep;
                else if (e.IsRoot)
                    data[i] = null;               // mini 流容器稍后统一构造
                else
                    data[i] = src.ReadStream(e);
            }

            // ---------- 2. 把小于阈值的流排进 mini 流 ----------
            MemoryStream miniMs = new MemoryStream();
            uint[] miniStart = new uint[src.Entries.Count];
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                miniStart[i] = CfbfFile.ENDOFCHAIN;
                if (!e.IsStream || data[i] == null || data[i].Length == 0) continue;
                if (data[i].Length >= cutoff) continue;

                // 按 mini 扇区对齐
                int need = ((data[i].Length + miniSize - 1) / miniSize) * miniSize;
                miniStart[i] = (uint)(miniMs.Length / miniSize);
                miniMs.Write(data[i], 0, data[i].Length);
                for (int k = data[i].Length; k < need; k++) miniMs.WriteByte(0);
            }
            byte[] miniStream = miniMs.ToArray();
            miniMs.Dispose();

            // mini FAT：每个 mini 扇区一个 4 字节项
            int miniSectors = miniStream.Length / miniSize;
            uint[] miniFat = new uint[Math.Max(miniSectors, 0)];
            for (int i = 0; i < miniFat.Length; i++) miniFat[i] = CfbfFile.FREESECT;
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                if (!e.IsStream || data[i] == null || data[i].Length == 0) continue;
                if (data[i].Length >= cutoff) continue;

                int cnt = (data[i].Length + miniSize - 1) / miniSize;
                uint start = miniStart[i];
                for (int k = 0; k < cnt; k++)
                    miniFat[start + k] = (k == cnt - 1) ? CfbfFile.ENDOFCHAIN : (uint)(start + k + 1);
            }

            // ---------- 3. 计算各段扇区数 ----------
            int wbSectors = 0;
            uint[] bigStart = new uint[src.Entries.Count];
            for (int i = 0; i < src.Entries.Count; i++) bigStart[i] = CfbfFile.ENDOFCHAIN;

            int bigCursor = 0;
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                if (!e.IsStream || data[i] == null || data[i].Length == 0) continue;
                if (data[i].Length < cutoff) continue;

                int cnt = (data[i].Length + sectorSize - 1) / sectorSize;
                bigStart[i] = (uint)bigCursor;
                bigCursor += cnt;
            }
            wbSectors = bigCursor;

            int miniStreamSectors = (miniStream.Length + sectorSize - 1) / sectorSize;
            int miniFatSectors = (miniFat.Length * 4 + sectorSize - 1) / sectorSize;
            int dirSectors = (src.Entries.Count * 128 + sectorSize - 1) / sectorSize;
            if (dirSectors < 1) dirSectors = 1;

            int dataSectors = wbSectors + miniStreamSectors + miniFatSectors + dirSectors;

            int fatSectors = 0;
            int difatSectors = 0;
            int perFatSector = sectorSize / 4;
            int perDifatSector = perFatSector - 1;
            for (int iter = 0; iter < 64; iter++)
            {
                int total = dataSectors + fatSectors + difatSectors;
                int needFat = (total * 4 + sectorSize - 1) / sectorSize;
                if (needFat < 1) needFat = 1;
                int needDifat = (needFat > 109)
                    ? (needFat - 109 + perDifatSector - 1) / perDifatSector
                    : 0;
                if (needFat == fatSectors && needDifat == difatSectors) break;
                fatSectors = needFat;
                difatSectors = needDifat;
            }

            // ---------- 4. 分配扇区号 ----------
            uint miniStreamStart = (uint)wbSectors;
            uint miniFatStart = (uint)(wbSectors + miniStreamSectors);
            uint dirStart = (uint)(wbSectors + miniStreamSectors + miniFatSectors);
            uint fatStart = (uint)(dataSectors);
            uint difatStart = (uint)(dataSectors + fatSectors);
            int totalSectors = dataSectors + fatSectors + difatSectors;

            // ---------- 5. 构造 FAT ----------
            uint[] fat = new uint[totalSectors];
            for (int i = 0; i < totalSectors; i++) fat[i] = CfbfFile.FREESECT;

            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                if (!e.IsStream || data[i] == null || data[i].Length == 0) continue;
                if (data[i].Length < cutoff) continue;

                int cnt = (data[i].Length + sectorSize - 1) / sectorSize;
                uint start = bigStart[i];
                for (int k = 0; k < cnt; k++)
                    fat[start + k] = (k == cnt - 1) ? CfbfFile.ENDOFCHAIN : (uint)(start + k + 1);
            }
            LinkChain(fat, miniStreamStart, miniStreamSectors);
            LinkChain(fat, miniFatStart, miniFatSectors);
            LinkChain(fat, dirStart, dirSectors);
            for (int k = 0; k < fatSectors; k++) fat[fatStart + k] = CfbfFile.FATSECT;
            for (int k = 0; k < difatSectors; k++) fat[difatStart + k] = CfbfFile.DIFSECT;

            // ---------- 6. 构造目录 ----------
            byte[] dir = new byte[dirSectors * sectorSize];
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                Buffer.BlockCopy(e.Raw, 0, dir, i * 128, 128);

                if (e.IsRoot)
                {
                    WriteU32(dir, i * 128 + 0x74, miniStreamSectors > 0 ? miniStreamStart : CfbfFile.ENDOFCHAIN);
                    WriteU64(dir, i * 128 + 0x78, miniStream.Length, src.MajorVersion);
                }
                else if (e.IsStream)
                {
                    byte[] d = data[i];
                    long size = (d == null) ? 0 : d.Length;
                    uint st = (size == 0) ? CfbfFile.ENDOFCHAIN
                             : (size < cutoff ? miniStart[i] : bigStart[i]);
                    WriteU32(dir, i * 128 + 0x74, st);
                    WriteU64(dir, i * 128 + 0x78, size, src.MajorVersion);
                }
            }

            // ---------- 7. 构造 mini FAT 与 FAT 的字节 ----------
            byte[] miniFatBytes = new byte[miniFatSectors * sectorSize];
            for (int i = 0; i < miniFat.Length; i++)
                WriteU32(miniFatBytes, i * 4, miniFat[i]);
            for (int i = miniFat.Length * 4; i < miniFatBytes.Length; i++) miniFatBytes[i] = 0xFF;

            byte[] fatBytes = new byte[fatSectors * sectorSize];
            for (int i = 0; i < fat.Length; i++)
                WriteU32(fatBytes, i * 4, fat[i]);
            for (int i = fat.Length * 4; i < fatBytes.Length; i++) fatBytes[i] = 0xFF;

            // ---------- 8. 头部 ----------
            byte[] header = new byte[512];
            Buffer.BlockCopy(src.Raw, 0, header, 0, 512);
            WriteU16(header, 0x1E, (ushort)TrailingZeros(sectorSize));
            WriteU16(header, 0x20, (ushort)TrailingZeros(miniSize));
            WriteU32(header, 0x28, src.MajorVersion >= 4 ? (uint)dirSectors : 0);
            WriteU32(header, 0x2C, (uint)fatSectors);
            WriteU32(header, 0x30, dirStart);
            WriteU32(header, 0x38, cutoff);
            WriteU32(header, 0x3C, miniFatSectors > 0 ? miniFatStart : CfbfFile.ENDOFCHAIN);
            WriteU32(header, 0x40, (uint)miniFatSectors);
            WriteU32(header, 0x44, difatSectors > 0 ? difatStart : CfbfFile.ENDOFCHAIN);
            WriteU32(header, 0x48, (uint)difatSectors);

            // 头部内嵌的 109 项 DIFAT
            for (int i = 0; i < 109; i++)
            {
                uint v = (i < fatSectors) ? (uint)(fatStart + i) : CfbfFile.FREESECT;
                WriteU32(header, 0x4C + i * 4, v);
            }

            // ---------- 9. 写文件 ----------
            // 头部占前 512 字节，相当于"第 -1 扇区"，所以扇区 n 的偏移是 (n+1)*sectorSize。
            // 当扇区大小就是 512 时，头部正好占满第一个扇区的位置，不会浪费。
            int headerPad = sectorSize - 512;
            if (headerPad < 0) headerPad = 0;

            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                fs.Write(header, 0, 512);
                if (headerPad > 0)
                {
                    byte[] pad = new byte[headerPad];
                    fs.Write(pad, 0, pad.Length);
                }

                WriteStreamAligned(fs, data, src, cutoff, bigStart, wbSectors, sectorSize);
                WriteBytesAligned(fs, miniStream, miniStreamSectors, sectorSize);
                WriteBytesAligned(fs, miniFatBytes, miniFatSectors, sectorSize);
                WriteBytesAligned(fs, dir, dirSectors, sectorSize);
                WriteBytesAligned(fs, fatBytes, fatSectors, sectorSize);

                if (difatSectors > 0)
                {
                    byte[] difatBytes = new byte[difatSectors * sectorSize];
                    for (int i = 0; i < difatBytes.Length; i++) difatBytes[i] = 0xFF;
                    int idx = 0;
                    uint cur = difatStart;
                    for (int s = 109; s < fatSectors; s++)
                    {
                        if (idx >= perDifatSector)
                        {
                            // 本扇区最后一项指向下一个 DIFAT 扇区
                            WriteU32(difatBytes, (int)((cur - difatStart) * sectorSize + idx * 4), cur + 1);
                            cur++;
                            idx = 0;
                        }
                        WriteU32(difatBytes, (int)((cur - difatStart) * sectorSize + idx * 4), (uint)(fatStart + s));
                        idx++;
                    }
                    WriteU32(difatBytes, (int)((cur - difatStart) * sectorSize + idx * 4), CfbfFile.ENDOFCHAIN);
                    fs.Write(difatBytes, 0, difatBytes.Length);
                }
            }
        }

        static void WriteStreamAligned(FileStream fs, byte[][] data, CfbfFile src, uint cutoff,
                                       uint[] bigStart, int wbSectors, int sectorSize)
        {
            byte[] buf = new byte[wbSectors * sectorSize];
            for (int i = 0; i < src.Entries.Count; i++)
            {
                CfbfEntry e = src.Entries[i];
                if (!e.IsStream || data[i] == null || data[i].Length == 0) continue;
                if (data[i].Length < cutoff) continue;

                int off = (int)bigStart[i] * sectorSize;
                Buffer.BlockCopy(data[i], 0, buf, off, data[i].Length);
            }
            fs.Write(buf, 0, buf.Length);
        }

        static void WriteBytesAligned(FileStream fs, byte[] b, int sectors, int sectorSize)
        {
            if (sectors <= 0) return;
            byte[] buf = new byte[sectors * sectorSize];
            Buffer.BlockCopy(b, 0, buf, 0, b.Length);
            fs.Write(buf, 0, buf.Length);
        }

        static void LinkChain(uint[] fat, uint start, int count)
        {
            for (int k = 0; k < count; k++)
                fat[start + k] = (k == count - 1) ? CfbfFile.ENDOFCHAIN : (uint)(start + k + 1);
        }

        static void WriteU16(byte[] b, int off, ushort v)
        {
            b[off] = (byte)(v & 0xFF);
            b[off + 1] = (byte)((v >> 8) & 0xFF);
        }

        static void WriteU32(byte[] b, int off, uint v)
        {
            b[off] = (byte)(v & 0xFF);
            b[off + 1] = (byte)((v >> 8) & 0xFF);
            b[off + 2] = (byte)((v >> 16) & 0xFF);
            b[off + 3] = (byte)((v >> 24) & 0xFF);
        }

        static void WriteU64(byte[] b, int off, long v, int majorVersion)
        {
            WriteU32(b, off, (uint)(v & 0xFFFFFFFF));
            if (majorVersion >= 4) WriteU32(b, off + 4, (uint)((v >> 32) & 0xFFFFFFFF));
            else WriteU32(b, off + 4, 0);
        }

        static int TrailingZeros(int v)
        {
            int n = 0;
            while (v > 1) { v >>= 1; n++; }
            return n;
        }
    }
}
