using System;
using System.IO;

namespace XlsxDoctor
{
    /// <summary>文件的真实容器格式。</summary>
    public enum FileKind
    {
        Unknown = 0,
        /// <summary>ZIP 包 —— .xlsx / .xlsm / .xltx / .xltm 等 OOXML 格式。</summary>
        Ooxml = 1,
        /// <summary>CFBF/OLE2 复合文档 —— .xls（BIFF5/BIFF8）/ .xlsb。</summary>
        Cfbf = 2
    }

    /// <summary>
    /// 按**文件内容**而不是扩展名判断格式。
    ///
    /// 现实里扩展名经常是错的：把 .xlsx 改名成 .xls 是极常见的操作（尤其是从邮件、
    /// 系统导出、或是"另存为"时选错类型）。Excel 自己就是按内容识别的，所以它照样能打开。
    /// 我们如果只看扩展名，就会对这类文件直接报"格式不支持"，白白拒绝掉一个完全能处理的文件。
    ///
    /// 只读文件头 8 个字节，成本可以忽略。
    /// </summary>
    public static class FormatSniffer
    {
        /// <summary>
        /// 是否启用 .xls（CFBF/BIFF8）清理。
        ///
        /// 引擎已经写好并通过测试（代码在 src\Xls\ 下，测试在 tests\XlsTests.ps1），
        /// 但当前版本暂时关闭入口 —— 改成 true 即可重新启用，不需要改别的地方。
        ///
        /// 关闭时 GUI 与 CLI 都会给出"暂不支持"的明确提示，而不是含糊的"格式错误"。
        /// </summary>
        public const bool XlsEnabled = false;

        public static FileKind Sniff(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    byte[] head = new byte[8];
                    int n = fs.Read(head, 0, 8);
                    if (n < 4) return FileKind.Unknown;

                    // ZIP: "PK" 0x03 0x04（也有 0x05 0x06 空档 / 0x07 0x08 分卷）
                    if (head[0] == 0x50 && head[1] == 0x4B
                        && (head[2] == 0x03 || head[2] == 0x05 || head[2] == 0x07))
                        return FileKind.Ooxml;

                    // CFBF: D0 CF 11 E0 A1 B1 1A E1
                    if (head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0)
                        return FileKind.Cfbf;

                    return FileKind.Unknown;
                }
            }
            catch (Exception)
            {
                return FileKind.Unknown;
            }
        }

        /// <summary>给报告用的中文描述。</summary>
        public static string Describe(FileKind k)
        {
            switch (k)
            {
                case FileKind.Ooxml: return "ZIP / OOXML 包（.xlsx 系列）";
                case FileKind.Cfbf: return "CFBF / OLE2 复合文档（.xls）";
                default: return "无法识别";
            }
        }

        /// <summary>
        /// 扩展名与真实内容是否矛盾。用于给用户一个明确的提示，
        /// 而不是让人对着"格式不支持"发懵。
        /// </summary>
        public static string ExtensionMismatch(string path, FileKind actual)
        {
            string ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext)) return null;
            ext = ext.ToLowerInvariant();

            bool claimsOoxml = (ext == ".xlsx" || ext == ".xlsm" || ext == ".xltx" || ext == ".xltm");
            bool claimsCfbf = (ext == ".xls" || ext == ".xlsb");

            if (actual == FileKind.Ooxml && claimsCfbf)
                return "扩展名是 " + ext + "，但文件内容其实是 OOXML（ZIP）包 —— 很可能是把 .xlsx 改名成了 .xls。已按真实格式处理。";
            if (actual == FileKind.Cfbf && claimsOoxml)
                return "扩展名是 " + ext + "，但文件内容其实是 CFBF/OLE2 复合文档 —— 很可能是把 .xls 改名成了 .xlsx。已按真实格式处理。";

            return null;
        }

        /// <summary>
        /// 这个文件能不能清理。不能则返回给用户看的中文原因，能则返回 null。
        ///
        /// 集中在这里判断，是为了让 GUI 和 CLI 说同样的话 —— 两边各写一套迟早会不一致。
        /// </summary>
        public static string RejectReason(string path)
        {
            FileKind kind = Sniff(path);

            if (kind == FileKind.Cfbf && !XlsEnabled)
            {
                return "这是 .xls 文件（老版 BIFF8 二进制格式），暂不支持。\r\n"
                     + "\r\n"
                     + "本工具目前只处理 OOXML 包：.xlsx / .xlsm / .xltx / .xltm。\r\n"
                     + "\r\n"
                     + "想清理的话，可以先用 Excel 打开它，「另存为」成 .xlsx，再用本工具清理。";
            }

            if (kind == FileKind.Unknown)
            {
                return "无法识别的文件格式，暂不支持。\r\n"
                     + "\r\n"
                     + "文件头既不是 ZIP/OOXML 包（PK），也不是 .xls 复合文档（D0 CF 11 E0）。\r\n"
                     + "它可能根本不是 Excel 文件，或者已经损坏。";
            }

            return null;
        }

        /// <summary>不支持的原因（单行版，给命令行用）。</summary>
        public static string RejectReasonOneLine(string path)
        {
            string r = RejectReason(path);
            if (r == null) return null;
            return r.Replace("\r\n", " ").Replace("  ", " ").Trim();
        }
    }
}
