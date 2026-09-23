using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace XlsxDoctor
{
    internal static class Program
    {
        internal const string Version = "1.1.0";

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(int dwProcessId);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        const int ATTACH_PARENT_PROCESS = -1;

        [STAThread]
        static int Main(string[] args)
        {
            // 分派规则：
            //   无参数                        -> 图形界面
            //   只给一个存在的文件、无任何选项 -> 图形界面并自动分析（拖放到 exe 上 / 双击打开）
            //   出现任何以 - 开头的选项        -> 命令行模式
            bool hasFlag = false;
            for (int i = 0; i < args.Length; i++)
                if (args[i].StartsWith("-", StringComparison.Ordinal)) { hasFlag = true; break; }

            bool useGui = !IsCliBuild()
                       && (args.Length == 0
                           || (!hasFlag && args.Length == 1 && File.Exists(args[0])));

            if (useGui)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm(args.Length == 1 ? args[0] : null));
                return 0;
            }

            return RunCli(args);
        }

        /// <summary>
        /// CFBF 重建器的无变换往返自检：读入 → 原样重写 → 读回逐字节比对。
        /// 容器层的改动最容易出错，这一步能把"写坏了"和"清理逻辑写坏了"分开定位。
        /// </summary>
        static int RepackSelfTest(string srcPath, string outPath)
        {
            if (outPath == null)
                outPath = Path.Combine(Path.GetTempPath(), "xlsdoctor-repack-test.xls");

            Xls.CfbfFile cf = Xls.CfbfFile.Read(srcPath);
            Xls.CfbfEntry wbe = cf.FindWorkbookEntry();
            if (wbe == null)
            {
                Console.Error.WriteLine("失败：目录里找不到 Workbook / Book 流。");
                return 1;
            }

            byte[] orig = cf.ReadStream(wbe);
            Dictionary<int, byte[]> rep = new Dictionary<int, byte[]>();
            rep[wbe.Index] = orig;
            Xls.CfbfWriter.Write(outPath, cf, rep);

            Xls.CfbfFile back = Xls.CfbfFile.Read(outPath);
            Xls.CfbfEntry wb2 = back.FindWorkbookEntry();
            byte[] round = back.ReadStream(wb2);

            bool same = (round.Length == orig.Length);
            int firstDiff = -1;
            if (same)
            {
                for (int k = 0; k < orig.Length; k++)
                    if (orig[k] != round[k]) { same = false; firstDiff = k; break; }
            }

            Console.WriteLine("CFBF 重建往返自检");
            Console.WriteLine("  源文件          : " + srcPath);
            Console.WriteLine("  源大小          : " + new FileInfo(srcPath).Length.ToString("N0", CultureInfo.InvariantCulture) + " 字节");
            Console.WriteLine("  输出            : " + outPath);
            Console.WriteLine("  输出大小        : " + new FileInfo(outPath).Length.ToString("N0", CultureInfo.InvariantCulture) + " 字节");
            Console.WriteLine("  Workbook 流长度 : " + orig.Length.ToString("N0", CultureInfo.InvariantCulture)
                              + " -> " + round.Length.ToString("N0", CultureInfo.InvariantCulture));
            Console.WriteLine("  目录项数        : " + cf.Entries.Count.ToString(CultureInfo.InvariantCulture)
                              + " -> " + back.Entries.Count.ToString(CultureInfo.InvariantCulture));

            // 每个流都逐一比对
            int checkedStreams = 0, badStreams = 0;
            for (int i = 0; i < cf.Entries.Count; i++)
            {
                Xls.CfbfEntry e = cf.Entries[i];
                if (!e.IsStream) continue;
                Xls.CfbfEntry e2 = back.FindEntry(e.Name);
                if (e2 == null) { badStreams++; Console.WriteLine("    流丢失: " + e.Name); continue; }
                byte[] a = cf.ReadStream(e);
                byte[] b = back.ReadStream(e2);
                checkedStreams++;
                if (a.Length != b.Length) { badStreams++; Console.WriteLine("    长度不符: " + e.Name); continue; }
                for (int k = 0; k < a.Length; k++)
                    if (a[k] != b[k]) { badStreams++; Console.WriteLine("    内容不符: " + e.Name + " @ " + k); break; }
            }

            Console.WriteLine("  流逐字节比对    : " + checkedStreams.ToString(CultureInfo.InvariantCulture) + " 个流，"
                              + (badStreams == 0 ? "全部一致" : badStreams.ToString(CultureInfo.InvariantCulture) + " 个不符"));
            if (!same)
                Console.WriteLine("  !! Workbook 流不一致，首个差异在偏移 " + firstDiff.ToString(CultureInfo.InvariantCulture));

            bool ok = same && badStreams == 0;
            Console.WriteLine("  结论            : " + (ok ? "通过" : "失败"));
            return ok ? 0 : 3;
        }

        /// <summary>XlsxDoctor-cli.exe 是控制台版，即使不带参数也应走命令行分支。</summary>
        static bool IsCliBuild()
        {
            try
            {
                string exe = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                return Path.GetFileNameWithoutExtension(exe).EndsWith("-cli", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception) { return false; }
        }

        // ============================ 命令行模式 ============================

        static int RunCli(string[] args)
        {
            EnsureConsole();

            List<string> files = new List<string>();
            CleanOptions opt = new CleanOptions();
            string explicitOut = null;
            bool xlsInfo = false;
            bool xlsRepack = false;

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                string low = a.ToLowerInvariant();

                if (low == "-h" || low == "--help" || low == "/?" || low == "help")
                {
                    PrintHelp();
                    return 0;
                }
                else if (low == "--cli") { /* 强制命令行模式，无额外动作 */ }
                else if (low == "--xls-info") xlsInfo = true;
                else if (low == "--xls-repack") xlsRepack = true;
                else if (low == "-n" || low == "--dry-run") opt.DryRun = true;
                else if (low == "-c" || low == "--conservative") opt.Conservative = true;
                else if (low == "-i" || low == "--in-place") opt.InPlace = true;
                else if (low == "-o" || low == "--out")
                {
                    if (i + 1 >= args.Length)
                    {
                        Console.Error.WriteLine("错误：-o/--out 后面缺少路径。");
                        return 2;
                    }
                    explicitOut = args[++i];
                }
                else if (a.StartsWith("-", StringComparison.Ordinal))
                {
                    Console.Error.WriteLine("错误：未知选项 " + a);
                    Console.Error.WriteLine();
                    PrintHelp();
                    return 2;
                }
                else
                {
                    files.Add(a);
                }
            }

            if (files.Count == 0)
            {
                Console.Error.WriteLine("错误：没有指定要处理的文件。");
                Console.Error.WriteLine();
                PrintHelp();
                return 2;
            }

            if (explicitOut != null && files.Count > 1)
            {
                Console.Error.WriteLine("错误：-o/--out 只能用于单个文件。");
                return 2;
            }

            Console.WriteLine();
            Console.WriteLine("XlsxDoctor " + Version + "  —  Excel 垃圾清理");
            Console.WriteLine();

            int exit = 0;
            for (int i = 0; i < files.Count; i++)
            {
                if (files.Count > 1)
                {
                    Console.WriteLine("============================================================");
                    Console.WriteLine("[" + (i + 1).ToString(CultureInfo.InvariantCulture) + "/"
                        + files.Count.ToString(CultureInfo.InvariantCulture) + "] " + files[i]);
                    Console.WriteLine("============================================================");
                }

                opt.OutputPath = explicitOut;
                opt.Log = null;

                try
                {
                    if (xlsRepack)
                    {
                        exit = RepackSelfTest(files[i], explicitOut);
                    }
                    else if (xlsInfo)
                    {
                        Console.Write(Xls.XlsDiagnostics.Describe(files[i]));
                    }
                    else
                    {
                        // 先拦不支持的格式：给明确原因，而不是让下层抛个含糊的异常
                        string reason = FormatSniffer.RejectReasonOneLine(files[i]);
                        if (reason != null)
                        {
                            Console.Error.WriteLine("跳过：" + reason);
                            exit = 1;
                        }
                        else
                        {
                            // 按内容分派：扩展名经常是错的，Excel 自己也只认内容
                            FileKind kind = FormatSniffer.Sniff(files[i]);
                            CleanReport rep = (kind == FileKind.Cfbf)
                                ? Xls.XlsCleaner.Run(files[i], opt)
                                : new XlsxCleaner().Run(files[i], opt);

                            rep.AddWarning(FormatSniffer.ExtensionMismatch(files[i], kind));

                            Console.Write(rep.ToText());

                            if (rep.ValidationErrors.Count > 0) exit = 3;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("失败：" + ex.Message);
                    exit = 1;
                }

                Console.WriteLine();
            }

            if (opt.DryRun) Console.WriteLine("（--dry-run：只分析，未写出任何文件）");
            return exit;
        }

        static void PrintHelp()
        {
            Console.WriteLine("XlsxDoctor " + Version + "  —  清理 Excel 内部累积的垃圾（失效定义名称 / 无人引用的外部链接）");
            Console.WriteLine();
            Console.WriteLine("用法：");
            Console.WriteLine("  XlsxDoctor.exe                        打开图形界面");
            Console.WriteLine("  XlsxDoctor.exe <文件>                打开图形界面并自动分析该文件");
            Console.WriteLine("  XlsxDoctor.exe <文件> [选项…]        命令行模式");
            Console.WriteLine();
            Console.WriteLine("选项：");
            Console.WriteLine("  -n, --dry-run          只分析，不写出任何文件");
            Console.WriteLine("  -c, --conservative     保守模式：只删明确失效的名称，保留所有值仍有效的名称");
            Console.WriteLine("  -i, --in-place         就地覆盖原文件（先自动备份为 <文件名>.bak）");
            Console.WriteLine("  -o, --out <路径>       指定输出文件（默认 <文件名> - 已清理.xlsx）");
            Console.WriteLine("      --cli              强制命令行模式");
            Console.WriteLine("      --xls-info         只打印 .xls 的内部结构（CFBF 头 / BIFF 记录统计 / 公式令牌自检），不改文件");
            Console.WriteLine("      --xls-repack       CFBF 重建器自检：读入 .xls 后原样重写，再读回逐字节比对");
            Console.WriteLine("  -h, --help             显示本帮助");
            Console.WriteLine();
            Console.WriteLine("支持格式：.xlsx / .xlsm / .xltx / .xltm（OOXML 包）");
            Console.WriteLine("          格式按**文件内容**识别，不看扩展名 —— 扩展名写错也能正确处理");
            Console.WriteLine("          注意：.xls（老版 BIFF8 二进制）暂不支持，会给出明确提示");
            Console.WriteLine();
            Console.WriteLine("示例：");
            Console.WriteLine("  XlsxDoctor.exe 报表.xlsx --dry-run");
            Console.WriteLine("  XlsxDoctor.exe 报表.xlsx -i");
            Console.WriteLine("  XlsxDoctor.exe *.xlsx -c");
            Console.WriteLine();
            Console.WriteLine("退出码：0=成功  1=出错  2=参数错误  3=输出包校验失败");
        }

        // ============================ 控制台附着 ============================

        /// <summary>
        /// XlsxDoctor.exe 编译为 winexe（无控制台，双击不闪黑框），从命令行调用时
        /// 需要把自己附着到调用者的控制台上；XlsxDoctor-cli.exe 编译为控制台程序，
        /// 本身就有控制台，这里直接跳过。
        /// </summary>
        static void EnsureConsole()
        {
            bool hasConsole = (GetConsoleWindow() != IntPtr.Zero);

            if (!hasConsole)
            {
                if (!AttachConsole(ATTACH_PARENT_PROCESS)) AllocConsole();

                Encoding enc;
                try { enc = Console.OutputEncoding; }
                catch (Exception) { enc = Encoding.Default; }

                try
                {
                    StreamWriter so = new StreamWriter(Console.OpenStandardOutput(), enc);
                    so.AutoFlush = true;
                    Console.SetOut(so);

                    StreamWriter se = new StreamWriter(Console.OpenStandardError(), enc);
                    se.AutoFlush = true;
                    Console.SetError(se);
                }
                catch (Exception) { }

                // 附着到父控制台时父进程的提示符已经打印过了，先换行避免串行
                Console.WriteLine();
            }
        }
    }
}
