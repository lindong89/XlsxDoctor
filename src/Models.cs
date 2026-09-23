using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace XlsxDoctor
{
    /// <summary>清理选项。</summary>
    public sealed class CleanOptions
    {
        /// <summary>保守模式：只删明确失效的名称，保留所有值仍有效的名称。</summary>
        public bool Conservative;

        /// <summary>只分析不写文件。</summary>
        public bool DryRun;

        /// <summary>就地覆盖源文件（先自动备份为 .bak）。</summary>
        public bool InPlace;

        /// <summary>显式指定输出路径；为空则自动生成 "xxx - 已清理.xlsx"。</summary>
        public string OutputPath;

        /// <summary>进度回调，可为 null。</summary>
        public Action<string> Log;
    }

    /// <summary>一次清理的分析与执行结果。</summary>
    public sealed class CleanReport
    {
        public string SourcePath = "";
        public string OutputPath = "";
        public string BackupPath = "";
        public long SourceSize;
        public long OutputSize;

        public int SheetCount;
        public int FormulaTextCount;

        public int NameTotal;
        public int NameUsed;
        public int NameSystem;
        public int NameBroken;
        public int NameUnreferenced;
        public bool Conservative;

        public int LinkTotal;
        public int LinkKeptByFormula;
        public int LinkKeptByName;
        public int LinkDropped;

        public int DroppedParts;
        public int WorkbookXmlBefore;
        public int WorkbookXmlAfter;

        public readonly List<string> KeptLinkTargets = new List<string>();
        public readonly List<string> Warnings = new List<string>();
        public readonly List<string> ValidationErrors = new List<string>();

        /// <summary>源文件本身就存在的包结构问题（非本工具引入），输出包原样继承。</summary>
        public readonly List<string> PreexistingIssues = new List<string>();

        /// <summary>
        /// 加一条警告。重复的不会重复加 —— 同一句话可能从不同层（入口分派、清理引擎）
        /// 各自加进来一次，靠调用方自觉去重迟早会漏。
        /// </summary>
        public void AddWarning(string w)
        {
            if (string.IsNullOrEmpty(w)) return;
            if (!Warnings.Contains(w)) Warnings.Add(w);
        }

        public bool Analyzed;
        public bool WroteFile;
        public bool NothingToDo;
        public bool ValidationPassed;
        public double ElapsedSeconds;

        /// <summary>体积压缩比（输出/源），0..1。</summary>
        public double SizeRatio
        {
            get { return SourceSize <= 0 ? 0.0 : (double)OutputSize / (double)SourceSize; }
        }

        // ================= .xls（CFBF/BIFF8）专用 =================
        // .xls 是二进制格式，术语和统计口径与 .xlsx 不同，报告分两套。

        /// <summary>true 表示这是一次 .xls 清理。</summary>
        public bool IsXls;

        public int XlsBiffRecordsBefore;
        public int XlsBiffRecordsAfter;
        public long XlsStreamBefore;
        public long XlsStreamAfter;
        public int XlsFormulaParsed;
        public int XlsFormulaRefs;
        public int XlsSupbookTotal;
        public int XlsSupbookKept;
        public int XlsExternSheetTotal;
        public int XlsExternSheetKept;
        public int XlsExternNameTotal;
        public int XlsExternNameDropped;
        public int XlsExternSheetToSelf;
        public int XlsVirtualSupbook = -1;

        public string ToXlsText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("================ 分析报告（.xls / BIFF8）================");
            sb.AppendLine("源文件        : " + SourcePath);
            sb.AppendLine("大小          : " + SourceSize.ToString("N0", CultureInfo.InvariantCulture) + " 字节");
            sb.AppendLine("工作表        : " + SheetCount.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("BIFF 记录     : " + XlsBiffRecordsBefore.ToString("N0", CultureInfo.InvariantCulture) + " 条逻辑记录");
            sb.AppendLine("工作簿流      : " + XlsStreamBefore.ToString("N0", CultureInfo.InvariantCulture) + " 字节");
            sb.AppendLine("公式          : " + XlsFormulaParsed.ToString("N0", CultureInfo.InvariantCulture)
                          + " 条全部解析成功，共 " + XlsFormulaRefs.ToString("N0", CultureInfo.InvariantCulture)
                          + " 处名称/外部表引用");
            sb.AppendLine();
            sb.AppendLine("定义名称      : 共 " + NameTotal.ToString("N0", CultureInfo.InvariantCulture) + " 条 NAME 记录");
            sb.AppendLine("  ├ 被公式引用，保留 : " + NameUsed.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("  ├ _xlnm 系统名保留 : " + NameSystem.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("  ├ 定义已失效，删除 : " + NameBroken.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("  └ " + (Conservative ? "无人引用，保留   : " : "无人引用，删除   : ")
                          + NameUnreferenced.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine();
            sb.AppendLine("外部表条目    : 共 " + XlsExternSheetTotal.ToString("N0", CultureInfo.InvariantCulture) + " 条 EXTERNSHEET");
            sb.AppendLine("  ├ 指向本工作簿，保留 : " + XlsExternSheetToSelf.ToString("N0", CultureInfo.InvariantCulture)
                          + "（三维引用基础，不是外部链接）");
            sb.AppendLine("  ├ 被公式引用，保留   : " + LinkKeptByFormula.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("  └ 无人引用，删除     : " + LinkDropped.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("外部工作簿    : 共 " + XlsSupbookTotal.ToString("N0", CultureInfo.InvariantCulture) + " 条 SUPBOOK"
                          + (XlsVirtualSupbook >= 0 ? "（含 1 条自引用）" : ""));
            sb.AppendLine("  ├ 保留             : " + XlsSupbookKept.ToString("N0", CultureInfo.InvariantCulture));
            sb.AppendLine("  └ 删除             : " + (XlsSupbookTotal - XlsSupbookKept).ToString("N0", CultureInfo.InvariantCulture));
            if (XlsExternNameTotal > 0)
            {
                sb.AppendLine("外部名称      : 共 " + XlsExternNameTotal.ToString("N0", CultureInfo.InvariantCulture) + " 条 EXTERNNAME");
                sb.AppendLine("  └ 删除             : " + XlsExternNameDropped.ToString("N0", CultureInfo.InvariantCulture));
            }
            sb.AppendLine("剔除记录      : " + DroppedParts.ToString("N0", CultureInfo.InvariantCulture) + " 条");

            if (KeptLinkTargets.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("!! 注意：有外部引用正在被公式使用，已保留。示例：");
                for (int i = 0; i < KeptLinkTargets.Count && i < 5; i++)
                    sb.AppendLine("     " + KeptLinkTargets[i]);
            }

            for (int i = 0; i < Warnings.Count; i++)
            {
                sb.AppendLine();
                sb.AppendLine("警告: " + Warnings[i]);
            }

            if (NothingToDo)
            {
                sb.AppendLine();
                sb.AppendLine("没有发现可清理的垃圾，未写出文件。");
                return sb.ToString();
            }

            if (WroteFile)
            {
                sb.AppendLine();
                sb.AppendLine("================ 结果 ================");
                sb.AppendLine("输出文件      : " + OutputPath);
                if (!string.IsNullOrEmpty(BackupPath)) sb.AppendLine("原文件备份    : " + BackupPath);
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "大小          : {0:N0} -> {1:N0} 字节  ({2:N1}%)", SourceSize, OutputSize, SizeRatio * 100.0));
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "工作簿流      : {0:N0} -> {1:N0} 字节", XlsStreamBefore, XlsStreamAfter));
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "BIFF 记录     : {0:N0} -> {1:N0} 条", XlsBiffRecordsBefore, XlsBiffRecordsAfter));
                sb.AppendLine("包校验        : " + (ValidationPassed
                    ? "通过（CFBF 可解析 / 记录结构完整 / 公式索引不越界 / BOUNDSHEET 偏移正确）"
                    : "失败"));
            }

            if (ValidationErrors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("校验失败（已删除输出文件，源文件未受影响）：");
                for (int i = 0; i < ValidationErrors.Count && i < 20; i++)
                    sb.AppendLine("  " + ValidationErrors[i]);
            }

            sb.AppendLine();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "耗时          : {0:N1} 秒", ElapsedSeconds));
            return sb.ToString();
        }

        public string ToText()
        {
            if (IsXls) return ToXlsText();

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("================ 分析报告 ================");
            sb.AppendLine("源文件        : " + SourcePath);
            sb.AppendLine("大小          : " + SourceSize.ToString("N0", CultureInfo.InvariantCulture) + " 字节");
            sb.AppendLine("工作表        : " + SheetCount.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("公式文本      : " + FormulaTextCount.ToString(CultureInfo.InvariantCulture)
                          + " 条（来自工作表/图表/条件格式/数据验证）");
            sb.AppendLine();
            sb.AppendLine("定义名称      : 共 " + NameTotal.ToString(CultureInfo.InvariantCulture) + " 个");
            sb.AppendLine("  ├ 被公式引用，保留 : " + NameUsed.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("  ├ _xlnm 系统名保留 : " + NameSystem.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("  ├ 值已失效，删除   : " + NameBroken.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("  └ " + (Conservative ? "无人引用，保留   : " : "无人引用，删除   : ")
                          + NameUnreferenced.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine();
            sb.AppendLine("外部链接      : 共 " + LinkTotal.ToString(CultureInfo.InvariantCulture) + " 个");
            sb.AppendLine("  ├ 公式直接引用，保留       : " + LinkKeptByFormula.ToString(CultureInfo.InvariantCulture));
            if (LinkKeptByName > 0)
                sb.AppendLine("  ├ 仅被保留的名称引用，保留 : " + LinkKeptByName.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("  └ 无人引用，删除           : " + LinkDropped.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("连带删除部件  : " + DroppedParts.ToString(CultureInfo.InvariantCulture) + " 个（含 .rels）");

            if (KeptLinkTargets.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("!! 注意：有外部链接正在被公式引用，已保留。");
                sb.AppendLine("   这些链接会继续指向外部文件；如需彻底断开请手工确认。示例：");
                int shown = 0;
                for (int i = 0; i < KeptLinkTargets.Count && shown < 5; i++)
                {
                    sb.AppendLine("     " + KeptLinkTargets[i]);
                    shown++;
                }
                if (KeptLinkTargets.Count > shown) sb.AppendLine("     ...");
            }

            for (int i = 0; i < Warnings.Count; i++)
            {
                sb.AppendLine();
                sb.AppendLine("警告: " + Warnings[i]);
            }

            if (PreexistingIssues.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("源文件自带的包结构问题（这些部件本工具未改动，输出包原样继承）：");
                for (int i = 0; i < PreexistingIssues.Count && i < 10; i++)
                    sb.AppendLine("  " + PreexistingIssues[i]);
                if (PreexistingIssues.Count > 10)
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "  ...（共 {0} 项）", PreexistingIssues.Count));
            }

            if (NothingToDo)
            {
                sb.AppendLine();
                sb.AppendLine("没有发现可清理的垃圾，未写出文件。");
                return sb.ToString();
            }

            if (WroteFile)
            {
                sb.AppendLine();
                sb.AppendLine("================ 结果 ================");
                sb.AppendLine("输出文件      : " + OutputPath);
                if (!string.IsNullOrEmpty(BackupPath)) sb.AppendLine("原文件备份    : " + BackupPath);
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "大小          : {0:N0} -> {1:N0} 字节  ({2:N1}%)", SourceSize, OutputSize, SizeRatio * 100.0));
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "workbook.xml  : {0:N0} -> {1:N0} 字符", WorkbookXmlBefore, WorkbookXmlAfter));
                if (!ValidationPassed)
                {
                    sb.AppendLine("包校验        : 失败");
                }
                else if (PreexistingIssues.Count > 0)
                {
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "包校验        : 通过（另有 {0} 项问题源文件自带，原样继承）", PreexistingIssues.Count));
                }
                else
                {
                    sb.AppendLine("包校验        : 通过（XML 可解析 / 关系可解析 / Content_Types 完整）");
                }
            }

            if (ValidationErrors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("校验失败（已删除输出文件，源文件未受影响）：");
                for (int i = 0; i < ValidationErrors.Count && i < 20; i++)
                    sb.AppendLine("  " + ValidationErrors[i]);
            }

            sb.AppendLine();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "耗时          : {0:N1} 秒", ElapsedSeconds));
            return sb.ToString();
        }
    }
}
