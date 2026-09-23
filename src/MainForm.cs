using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XlsxDoctor
{
    public sealed class MainForm : Form
    {
        // ---------------- 控件 ----------------
        Panel _topPanel;
        Panel _dropZone;
        Label _dropHint;
        Label _titleLabel;
        Label _fileLabel;
        TextBox _pathBox;
        Button _browseBtn;
        CheckBox _conservativeChk;
        CheckBox _inPlaceChk;
        Button _analyzeBtn;
        Button _cleanBtn;
        Button _openFolderBtn;
        TextBox _reportBox;
        Label _statusLabel;

        string _currentFile;
        CleanReport _lastReport;
        bool _busy;

        public MainForm() : this(null) { }

        public MainForm(string initialFile)
        {
            BuildUi();
            ApplyIcon();
            if (!string.IsNullOrEmpty(initialFile)) _pendingFile = initialFile;
        }

        /// <summary>
        /// 窗口/任务栏图标。exe 本身已经用 /win32icon 内嵌了图标，这里直接把它取出来用，
        /// 不需要再把 .ico 当资源嵌一遍 —— 少一份会不同步的副本。
        /// </summary>
        void ApplyIcon()
        {
            try
            {
                Icon exeIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (exeIcon != null) Icon = exeIcon;
            }
            catch (Exception)
            {
                // 取不到就退回系统默认图标，不值得为此让程序启动失败
            }
        }

        string _pendingFile;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (!string.IsNullOrEmpty(_pendingFile))
            {
                string f = _pendingFile;
                _pendingFile = null;
                if (RejectIfUnsupported(f)) return;
                LoadFile(f);
            }
            else
            {
                SetStatus("把 .xlsx / .xlsm 文件拖进上面的方框，或点击「选择文件」开始。");
            }
        }

        // ============================ 界面构建 ============================

        void BuildUi()
        {
            Text = "XlsxDoctor — Excel 垃圾清理（定义名称 / 外部链接）";
            ClientSize = new Size(920, 680);
            MinimumSize = new Size(760, 520);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(246, 247, 249);
            AllowDrop = true;

            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 3;
            // 必须显式声明列宽：不设 ColumnStyles 时该列会退化成 AutoSize，
            // 被内容撑到 1600px 以上，锚定 Right 的控件会被挤出窗口。
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 206F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));
            root.BackColor = Color.FromArgb(246, 247, 249);

            // ---------- 顶部 ----------
            _topPanel = new Panel();
            _topPanel.Dock = DockStyle.Fill;
            _topPanel.Margin = new Padding(0);
            _topPanel.AutoSize = false;
            // 关键：必须先把宽度设成真实值再往里加锚定控件。Panel 默认宽 200px，
            // 若此时加锚定控件，右边距会算成 200-16-888 = -704，等 Dock=Fill 把
            // 本面板拉到 920px 后，控件会整体涨 720px 并跑出窗口。
            _topPanel.Size = new Size(920, 206);

            _titleLabel = new Label();
            _titleLabel.Text = "XlsxDoctor";
            _titleLabel.Font = new Font("Microsoft YaHei UI", 15F, FontStyle.Bold);
            _titleLabel.ForeColor = Color.FromArgb(32, 44, 66);
            _titleLabel.AutoSize = true;
            _titleLabel.Location = new Point(16, 12);

            Label sub = new Label();
            sub.Text = "清理 .xlsx 内部累积的垃圾：失效的定义名称 + 无人引用的外部链接";
            sub.ForeColor = Color.FromArgb(110, 120, 135);
            sub.AutoSize = true;
            sub.Location = new Point(140, 22);

            _dropZone = new Panel();
            _dropZone.Location = new Point(16, 46);
            _dropZone.Size = new Size(888, 50);
            _dropZone.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _dropZone.BorderStyle = BorderStyle.FixedSingle;
            _dropZone.BackColor = Color.White;
            _dropZone.AllowDrop = true;
            _dropZone.DragEnter += OnDragEnter;
            _dropZone.DragDrop += OnDragDrop;
            _dropZone.Cursor = Cursors.Hand;
            _dropZone.Click += delegate { Browse(); };

            _dropHint = new Label();
            _dropHint.Text = "把 .xlsx / .xlsm 拖到这里  ·  或点击此处选择文件";
            _dropHint.Dock = DockStyle.Fill;
            _dropHint.TextAlign = ContentAlignment.MiddleCenter;
            _dropHint.ForeColor = Color.FromArgb(120, 130, 145);
            _dropHint.Cursor = Cursors.Hand;
            _dropHint.Click += delegate { Browse(); };
            _dropZone.Controls.Add(_dropHint);

            _fileLabel = new Label();
            _fileLabel.Text = "文件";
            _fileLabel.AutoSize = true;
            _fileLabel.Location = new Point(18, 108);
            _fileLabel.ForeColor = Color.FromArgb(80, 90, 105);

            // 路径行：文本框 + 按钮放进一个 TableLayoutPanel——
            // 文本框吃掉剩余宽度、按钮固定 132px，无论窗口多宽 / 什么 DPI 都不会互相覆盖。
            // （直接锚定时，WinForms 在某些布局时序/缩放下会把锚定边距算偏，
            //   文本框会涨宽盖住按钮。列宽百分比 + 固定列则结构上不可能重叠。）
            TableLayoutPanel pathRow = new TableLayoutPanel();
            pathRow.Location = new Point(62, 102);
            pathRow.Size = new Size(818, 32);
            pathRow.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            pathRow.ColumnCount = 2;
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132F));
            pathRow.Margin = new Padding(0);
            pathRow.BackColor = Color.Transparent;

            _pathBox = new TextBox();
            _pathBox.Dock = DockStyle.Fill;
            _pathBox.Margin = new Padding(0, 3, 8, 3);
            _pathBox.ReadOnly = true;
            _pathBox.BackColor = Color.White;

            _browseBtn = new Button();
            _browseBtn.Text = "选择文件…";
            _browseBtn.Dock = DockStyle.Fill;
            _browseBtn.Margin = new Padding(0, 1, 0, 1);
            _browseBtn.Click += delegate { Browse(); };

            pathRow.Controls.Add(_pathBox, 0, 0);
            pathRow.Controls.Add(_browseBtn, 1, 0);

            _conservativeChk = new CheckBox();
            _conservativeChk.Text = "保守模式（只删明确失效的名称）";
            _conservativeChk.AutoSize = true;
            _conservativeChk.Location = new Point(62, 138);
            _conservativeChk.ForeColor = Color.FromArgb(80, 90, 105);

            _inPlaceChk = new CheckBox();
            _inPlaceChk.Text = "就地覆盖原文件（自动备份 .bak）";
            _inPlaceChk.AutoSize = true;
            _inPlaceChk.Location = new Point(330, 138);
            _inPlaceChk.ForeColor = Color.FromArgb(80, 90, 105);

            _analyzeBtn = new Button();
            _analyzeBtn.Text = "1  仅分析";
            _analyzeBtn.Location = new Point(60, 166);
            _analyzeBtn.Size = new Size(112, 30);
            _analyzeBtn.Enabled = false;
            _analyzeBtn.Click += delegate { StartAnalyze(); };

            _cleanBtn = new Button();
            _cleanBtn.Text = "2  修复并保存";
            _cleanBtn.Location = new Point(180, 166);
            _cleanBtn.Size = new Size(140, 30);
            _cleanBtn.Enabled = false;
            _cleanBtn.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            _cleanBtn.Click += delegate { StartClean(); };

            _openFolderBtn = new Button();
            _openFolderBtn.Text = "打开所在文件夹";
            _openFolderBtn.Location = new Point(328, 166);
            _openFolderBtn.Size = new Size(140, 30);
            _openFolderBtn.Enabled = false;
            _openFolderBtn.Click += delegate { OpenFolder(); };

            _topPanel.Controls.Add(_titleLabel);
            _topPanel.Controls.Add(sub);
            _topPanel.Controls.Add(_dropZone);
            _topPanel.Controls.Add(_fileLabel);
            _topPanel.Controls.Add(pathRow);
            _topPanel.Controls.Add(_conservativeChk);
            _topPanel.Controls.Add(_inPlaceChk);
            _topPanel.Controls.Add(_analyzeBtn);
            _topPanel.Controls.Add(_cleanBtn);
            _topPanel.Controls.Add(_openFolderBtn);

            // ---------- 报告区 ----------
            _reportBox = new TextBox();
            _reportBox.Dock = DockStyle.Fill;
            _reportBox.Multiline = true;
            _reportBox.ReadOnly = true;
            _reportBox.ScrollBars = ScrollBars.Both;
            _reportBox.WordWrap = false;
            _reportBox.BackColor = Color.White;
            _reportBox.BorderStyle = BorderStyle.FixedSingle;
            _reportBox.Font = new Font("Consolas", 9F);
            _reportBox.Margin = new Padding(16, 4, 16, 4);
            _reportBox.Text = "尚未选择文件。\r\n";

            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("复制", null, delegate { if (_reportBox.SelectionLength > 0) _reportBox.Copy(); else { _reportBox.SelectAll(); _reportBox.Copy(); } });
            menu.Items.Add("全选", null, delegate { _reportBox.SelectAll(); });
            _reportBox.ContextMenuStrip = menu;

            // ---------- 状态栏（左：状态文字，右：版本号） ----------
            TableLayoutPanel statusRow = new TableLayoutPanel();
            statusRow.Dock = DockStyle.Fill;
            statusRow.ColumnCount = 2;
            statusRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            statusRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72F));
            statusRow.Margin = new Padding(0);
            statusRow.BackColor = Color.FromArgb(236, 238, 242);

            _statusLabel = new Label();
            _statusLabel.Dock = DockStyle.Fill;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statusLabel.Padding = new Padding(18, 0, 0, 0);
            _statusLabel.ForeColor = Color.FromArgb(90, 100, 115);
            _statusLabel.BackColor = Color.FromArgb(236, 238, 242);
            _statusLabel.Text = "就绪";

            Label versionLabel = new Label();
            versionLabel.Text = "v" + Program.Version;
            versionLabel.Dock = DockStyle.Fill;
            versionLabel.TextAlign = ContentAlignment.MiddleRight;
            versionLabel.Padding = new Padding(0, 0, 12, 0);
            versionLabel.ForeColor = Color.FromArgb(130, 140, 155);
            versionLabel.BackColor = Color.FromArgb(236, 238, 242);

            statusRow.Controls.Add(_statusLabel, 0, 0);
            statusRow.Controls.Add(versionLabel, 1, 0);

            root.Controls.Add(_topPanel, 0, 0);
            root.Controls.Add(_reportBox, 0, 1);
            root.Controls.Add(statusRow, 0, 2);
            Controls.Add(root);
        }

        // ============================ 拖放 ============================

        static bool IsSupported(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string e = Path.GetExtension(path).ToLowerInvariant();
            if (e == ".xlsb") return false;   // BIFF12，和 BIFF8 是两套格式，不支持
            if (e == ".xlsx" || e == ".xlsm" || e == ".xltx" || e == ".xltm") return true;
            if (e == ".xls") return FormatSniffer.XlsEnabled;

            // 扩展名不认识时，按文件内容再判一次 —— 现实里扩展名经常是错的
            FileKind k = FormatSniffer.Sniff(path);
            if (k == FileKind.Ooxml) return true;
            if (k == FileKind.Cfbf) return FormatSniffer.XlsEnabled;
            return false;
        }

        /// <summary>按文件真实内容分派到对应的清理引擎。</summary>
        static CleanReport RunCleaner(string file, CleanOptions opt)
        {
            // 兜底：入口处已经拦过，这里再拦一次，防止将来从别的路径调进来时静默走错引擎
            string reason = FormatSniffer.RejectReason(file);
            if (reason != null) throw new NotSupportedException(reason);

            return (FormatSniffer.Sniff(file) == FileKind.Cfbf)
                ? Xls.XlsCleaner.Run(file, opt)
                : new XlsxCleaner().Run(file, opt);
        }

        /// <summary>把「暂不支持」的原因弹给用户看。返回 false 表示应该中止这次操作。</summary>
        bool RejectIfUnsupported(string path)
        {
            string reason = FormatSniffer.RejectReason(path);
            if (reason == null) return false;

            MessageBox.Show(this, reason, "暂不支持", MessageBoxButtons.OK, MessageBoxIcon.Information);
            SetStatus("暂不支持这个文件格式。");
            return true;
        }

        void OnDragEnter(object sender, DragEventArgs e)
        {
            if (_busy) { e.Effect = DragDropEffects.None; return; }
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effect = DragDropEffects.None; return; }
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effect = (files != null && files.Length > 0 && IsSupported(files[0]))
                ? DragDropEffects.Copy : DragDropEffects.None;
        }

        void OnDragDrop(object sender, DragEventArgs e)
        {
            if (_busy) return;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            if (RejectIfUnsupported(files[0])) return;
            LoadFile(files[0]);
        }

        void Browse()
        {
            if (_busy) return;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "选择要清理的 Excel 文件";
            // .xls 暂不支持，就不列进筛选器（列了反而误导）；仍保留"所有文件"以便用户
            // 选中扩展名写错的文件 —— 格式最终按内容判定
            dlg.Filter = "Excel 工作簿 (*.xlsx;*.xlsm;*.xltx;*.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm"
                       + "|所有文件 (*.*)|*.*";
            dlg.CheckFileExists = true;
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                if (RejectIfUnsupported(dlg.FileName)) return;
                LoadFile(dlg.FileName);
            }
        }

        // ============================ 流程 ============================

        void LoadFile(string path)
        {
            _currentFile = path;
            _lastReport = null;
            _pathBox.Text = path;
            _dropHint.Text = Path.GetFileName(path);
            _dropHint.ForeColor = Color.FromArgb(32, 44, 66);
            _openFolderBtn.Enabled = false;
            StartAnalyze();
        }

        async void StartAnalyze()
        {
            if (_busy || string.IsNullOrEmpty(_currentFile)) return;
            SetBusy(true);
            SetStatus("正在分析…");
            _reportBox.Text = "正在分析，请稍候…\r\n\r\n（超大文件可能需要几十秒）\r\n";

            string file = _currentFile;
            CleanOptions opt = new CleanOptions();
            opt.DryRun = true;
            opt.Conservative = _conservativeChk.Checked;
            opt.Log = MakeLogger();

            try
            {
                CleanReport rep = await Task.Run(delegate { return RunCleaner(file, opt); });
                _lastReport = rep;
                ShowReport(rep);
                _cleanBtn.Enabled = !rep.NothingToDo;
                SetStatus(rep.NothingToDo
                    ? "分析完成：没有发现可清理的垃圾。"
                    : string.Format("分析完成：可删除 {0:N0} 个定义名称、{1:N0} 个{2}。点「修复并保存」执行。",
                        rep.NameBroken + rep.NameUnreferenced, rep.LinkDropped,
                        rep.IsXls ? "外部表条目" : "外部链接"));
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        async void StartClean()
        {
            if (_busy || string.IsNullOrEmpty(_currentFile)) return;
            SetBusy(true);
            SetStatus("正在修复…");
            _reportBox.Text = "正在修复并保存，请稍候…\r\n\r\n（超大文件可能需要几十秒）\r\n";

            string file = _currentFile;
            CleanOptions opt = new CleanOptions();
            opt.Conservative = _conservativeChk.Checked;
            opt.InPlace = _inPlaceChk.Checked;
            opt.Log = MakeLogger();

            try
            {
                CleanReport rep = await Task.Run(delegate { return RunCleaner(file, opt); });
                _lastReport = rep;
                ShowReport(rep);

                if (rep.ValidationErrors.Count > 0)
                {
                    _openFolderBtn.Enabled = false;
                    SetStatus("输出包校验失败，已丢弃输出文件（源文件未受影响）。");
                    MessageBox.Show(this,
                        "输出包校验未通过，为安全起见已删除输出文件。\r\n源文件未受任何影响。\r\n\r\n详见报告。",
                        "校验失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (rep.NothingToDo)
                {
                    _cleanBtn.Enabled = false;
                    SetStatus("没有发现可清理的垃圾，未写出文件。");
                }
                else if (rep.WroteFile)
                {
                    _cleanBtn.Enabled = false;
                    _openFolderBtn.Enabled = true;
                    SetStatus(string.Format("完成：{0:N0} → {1:N0} 字节（{2:N1}%）。{3}",
                        rep.SourceSize, rep.OutputSize, rep.SizeRatio * 100.0,
                        rep.BackupPath.Length > 0 ? "原文件已备份为 .bak" : ""));
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        Action<string> MakeLogger()
        {
            return delegate(string msg)
            {
                if (IsDisposed) return;
                try
                {
                    BeginInvoke((Action)delegate { if (!IsDisposed) _statusLabel.Text = msg; });
                }
                catch (Exception) { }
            };
        }

        void ShowReport(CleanReport rep)
        {
            _reportBox.Text = rep.ToText();
            _reportBox.Select(0, 0);
            _reportBox.ScrollToCaret();
        }

        void ShowError(Exception ex)
        {
            SetStatus("出错：" + ex.Message);
            _reportBox.Text = "出错：\r\n\r\n" + ex.Message + "\r\n\r\n" + ex.GetType().FullName
                + "\r\n\r\n" + (ex.StackTrace == null ? "" : ex.StackTrace);
            MessageBox.Show(this, ex.Message, "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        void OpenFolder()
        {
            string target = null;
            if (_lastReport != null && _lastReport.WroteFile && File.Exists(_lastReport.OutputPath))
                target = _lastReport.OutputPath;
            else if (!string.IsNullOrEmpty(_currentFile) && File.Exists(_currentFile))
                target = _currentFile;

            if (target == null) return;
            try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + target + "\""); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "无法打开文件夹", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        void SetBusy(bool busy)
        {
            _busy = busy;
            _browseBtn.Enabled = !busy;
            _analyzeBtn.Enabled = !busy && !string.IsNullOrEmpty(_currentFile);
            if (busy) _cleanBtn.Enabled = false;
            _conservativeChk.Enabled = !busy;
            _inPlaceChk.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        void SetStatus(string text)
        {
            _statusLabel.Text = text;
        }
    }
}
