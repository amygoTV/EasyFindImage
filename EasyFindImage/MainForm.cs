using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace EasyFindImage
{
    public partial class MainForm : Form
    {
        // ==================== Tab 1：表格汇总与筛选 ====================
        private TabControl tabControl;
        private TabPage tabSummary;
        private TabPage tabCopy;
        private TabPage tabSettings;

        private GroupBox grpSource;
        private TextBox txtSourceFolder;
        private Button btnBrowseSource;
        private Button btnSummary;

        private GroupBox grpResult;
        private DataGridView dgvSummary;

        private GroupBox grpFilter;
        private Label lblFilterColumn;
        private ComboBox cbxFilterColumn;
        private Label lblFilterValue;
        private TextBox txtFilterValue;
        private Button btnFilter;
        private Button btnClearFilter;
        private Button btnExport;

        private GroupBox grpClean;
        private Label lblCleanColumn;
        private ComboBox cbxCleanColumn;
        private Label lblCleanKeywords;
        private TextBox txtCleanKeywords;
        private Button btnAddCleanCondition;
        private ListBox lstCleanConditions;
        private Button btnRemoveCleanCondition;
        private Button btnClearCleanConditions;
        private Button btnExecuteClean;
        private CheckBox cbxRemoveDuplicateSN;
        private CheckBox cbxKeepOnlyCommonCols;

        private Label lblStatus;
        private ProgressBar pgbStatus;

        // ==================== Tab 2：SN 查找图片 ====================
        private GroupBox grpSnSource;
        private RadioButton rdoFromExcel;
        private RadioButton rdoFromTxt;
        private TextBox txtSnSource;
        private Button btnBrowseSnSource;

        private GroupBox grpImageRoot;
        private TextBox txtImageRoot;
        private Button btnBrowseImageRoot;

        private Label lblSearchScope;
        private RadioButton rdoSpecifiedDir;
        private RadioButton rdoAllDrives;

        private GroupBox grpTarget;
        private TextBox txtTargetDir;
        private Button btnBrowseTarget;

        private Button btnStartCopy;
        private ProgressBar pgbCopy;
        private Label lblCopyStatus;
        private ListBox lstLog;

        private DataTable _summaryTable = new DataTable();
        private BindingSource _bindingSource = new BindingSource();

        private Bitmap _appIconBitmap;

        public MainForm()
        {
            this.Text = "EasyFindImage - 表格汇总与 SN 图片查找工具";
            this.Size = new Size(960, 760);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(880, 680);
            this.Icon = CreateAppIcon();

            InitializeUI();

            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void InitializeUI()
        {
            tabControl = new TabControl { Dock = DockStyle.Fill };

            tabSummary = new TabPage { Text = "表格汇总与筛选" };
            BuildSummaryTab(tabSummary);
            tabControl.TabPages.Add(tabSummary);

            tabCopy = new TabPage { Text = "按 SN 查找图片" };
            BuildCopyTab(tabCopy);
            tabControl.TabPages.Add(tabCopy);

            tabSettings = new TabPage { Text = "一键电脑设置" };
            BuildSettingsTab(tabSettings);
            tabControl.TabPages.Add(tabSettings);

            this.Controls.Add(tabControl);
        }

        private void BuildSummaryTab(TabPage tab)
        {
            int y = 15;

            // 源文件夹
            grpSource = new GroupBox { Text = "1. 选择目标文件夹", Left = 12, Top = y, Width = 900, Height = 70 };
            txtSourceFolder = new TextBox { Left = 15, Top = 22, Width = 620, ReadOnly = true };
            btnBrowseSource = new Button { Text = "浏览...", Left = 645, Top = 20, Width = 80, Height = 26 };
            btnSummary = new Button { Text = "汇总所有表格", Left = 735, Top = 20, Width = 110, Height = 26 };
            btnSummary.Enabled = false;
            grpSource.Controls.Add(txtSourceFolder);
            grpSource.Controls.Add(btnBrowseSource);
            grpSource.Controls.Add(btnSummary);
            tab.Controls.Add(grpSource);

            btnBrowseSource.Click += BtnBrowseSource_Click;
            btnSummary.Click += BtnSummary_Click;

            y += 85;

            // 筛选（可单独执行，列名自动从表头生成）
            grpFilter = new GroupBox { Text = "3. 筛选数据行", Left = 12, Top = y, Width = 900, Height = 70 };
            lblFilterColumn = new Label { Text = "列名:", Left = 15, Top = 26, Width = 40 };
            cbxFilterColumn = new ComboBox { Left = 55, Top = 22, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            lblFilterValue = new Label { Text = "关键字:", Left = 215, Top = 26, Width = 55 };
            txtFilterValue = new TextBox { Left = 270, Top = 22, Width = 180 };
            btnFilter = new Button { Text = "筛选", Left = 460, Top = 20, Width = 70, Height = 26 };
            btnClearFilter = new Button { Text = "清除筛选", Left = 540, Top = 20, Width = 80, Height = 26 };
            btnExport = new Button { Text = "导出结果", Left = 630, Top = 20, Width = 90, Height = 26 };
            grpFilter.Controls.Add(lblFilterColumn);
            grpFilter.Controls.Add(cbxFilterColumn);
            grpFilter.Controls.Add(lblFilterValue);
            grpFilter.Controls.Add(txtFilterValue);
            grpFilter.Controls.Add(btnFilter);
            grpFilter.Controls.Add(btnClearFilter);
            grpFilter.Controls.Add(btnExport);
            tab.Controls.Add(grpFilter);

            btnFilter.Click += BtnFilter_Click;
            btnClearFilter.Click += BtnClearFilter_Click;
            btnExport.Click += BtnExport_Click;

            y += 80;

            // 数据清理条件（可单独执行，关键字用逗号分隔表示"或"关系）
            grpClean = new GroupBox { Text = "4. 数据清理条件（同一列多个关键字用逗号分隔，任一匹配即删除；关键字留空表示该列为空白）", Left = 12, Top = y, Width = 900, Height = 125 };
            lblCleanColumn = new Label { Text = "列名:", Left = 15, Top = 24, Width = 40 };
            cbxCleanColumn = new ComboBox { Left = 55, Top = 20, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            lblCleanKeywords = new Label { Text = "关键字:", Left = 215, Top = 24, Width = 55 };
            txtCleanKeywords = new TextBox { Left = 270, Top = 20, Width = 260 };
            btnAddCleanCondition = new Button { Text = "添加条件", Left = 540, Top = 18, Width = 80, Height = 26 };
            lstCleanConditions = new ListBox { Left = 15, Top = 52, Width = 600, Height = 80 };
            btnRemoveCleanCondition = new Button { Text = "移除选中", Left = 625, Top = 52, Width = 80, Height = 24 };
            btnClearCleanConditions = new Button { Text = "清空条件", Left = 715, Top = 52, Width = 80, Height = 24 };
            cbxRemoveDuplicateSN = new CheckBox { Text = "删除重复SN行", Left = 660, Top = 20, Width = 120 };
            cbxKeepOnlyCommonCols = new CheckBox { Text = "只保留共有列", Left = 790, Top = 20, Width = 100 };
            btnExecuteClean = new Button { Text = "执行清理", Left = 800, Top = 52, Width = 80, Height = 26 };
            grpClean.Controls.Add(lblCleanColumn);
            grpClean.Controls.Add(cbxCleanColumn);
            grpClean.Controls.Add(lblCleanKeywords);
            grpClean.Controls.Add(txtCleanKeywords);
            grpClean.Controls.Add(btnAddCleanCondition);
            grpClean.Controls.Add(lstCleanConditions);
            grpClean.Controls.Add(btnRemoveCleanCondition);
            grpClean.Controls.Add(btnClearCleanConditions);
            grpClean.Controls.Add(cbxRemoveDuplicateSN);
            grpClean.Controls.Add(cbxKeepOnlyCommonCols);
            grpClean.Controls.Add(btnExecuteClean);
            tab.Controls.Add(grpClean);

            btnAddCleanCondition.Click += BtnAddCleanCondition_Click;
            btnRemoveCleanCondition.Click += BtnRemoveCleanCondition_Click;
            btnClearCleanConditions.Click += BtnClearCleanConditions_Click;
            btnExecuteClean.Click += BtnExecuteClean_Click;

            y += 135;

            // 汇总结果
            grpResult = new GroupBox { Text = "2. 汇总结果", Left = 12, Top = y, Width = 900, Height = 250 };
            dgvSummary = new DataGridView
            {
                Left = 15,
                Top = 22,
                Width = 870,
                Height = 200,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            grpResult.Controls.Add(dgvSummary);
            tab.Controls.Add(grpResult);

            y += 265;

            // 状态
            lblStatus = new Label { Left = 12, Top = y, Width = 300, Height = 20, Text = "就绪" };
            pgbStatus = new ProgressBar { Left = 12, Top = y + 2, Width = 876, Height = 14 };
            tab.Controls.Add(lblStatus);
            tab.Controls.Add(pgbStatus);
        }

        private void BuildCopyTab(TabPage tab)
        {
            int y = 15;

            // SN 来源
            grpSnSource = new GroupBox { Text = "1. SN 来源（表格或TXT）", Left = 12, Top = y, Width = 900, Height = 75 };
            rdoFromExcel = new RadioButton { Text = "来自表格", Left = 15, Top = 22, Width = 80, Checked = true };
            rdoFromTxt = new RadioButton { Text = "来自TXT文本", Left = 105, Top = 22, Width = 110 };
            txtSnSource = new TextBox { Left = 225, Top = 20, Width = 500, ReadOnly = true };
            btnBrowseSnSource = new Button { Text = "浏览...", Left = 735, Top = 18, Width = 80, Height = 26 };
            grpSnSource.Controls.Add(rdoFromExcel);
            grpSnSource.Controls.Add(rdoFromTxt);
            grpSnSource.Controls.Add(txtSnSource);
            grpSnSource.Controls.Add(btnBrowseSnSource);
            tab.Controls.Add(grpSnSource);
            btnBrowseSnSource.Click += BtnBrowseSnSource_Click;

            y += 90;

            // 图片根目录
            grpImageRoot = new GroupBox { Text = "2. 图片根目录 / 查找范围", Left = 12, Top = y, Width = 900, Height = 100 };
            txtImageRoot = new TextBox { Left = 15, Top = 22, Width = 560, ReadOnly = true };
            btnBrowseImageRoot = new Button { Text = "浏览...", Left = 585, Top = 20, Width = 80, Height = 26 };
            lblSearchScope = new Label { Text = "查找范围:", Left = 15, Top = 56, Width = 70 };
            rdoSpecifiedDir = new RadioButton { Text = "指定目录查找", Left = 90, Top = 52, Width = 120, Checked = true };
            rdoAllDrives = new RadioButton { Text = "全盘递归查找（从磁盘根目录开始，较慢）", Left = 220, Top = 52, Width = 300 };
            grpImageRoot.Controls.Add(txtImageRoot);
            grpImageRoot.Controls.Add(btnBrowseImageRoot);
            grpImageRoot.Controls.Add(lblSearchScope);
            grpImageRoot.Controls.Add(rdoSpecifiedDir);
            grpImageRoot.Controls.Add(rdoAllDrives);
            tab.Controls.Add(grpImageRoot);
            btnBrowseImageRoot.Click += BtnBrowseImageRoot_Click;

            y += 110;

            // 目标路径
            grpTarget = new GroupBox { Text = "3. 复制到目标路径", Left = 12, Top = y, Width = 900, Height = 65 };
            txtTargetDir = new TextBox { Left = 15, Top = 22, Width = 620, ReadOnly = true };
            btnBrowseTarget = new Button { Text = "浏览...", Left = 645, Top = 20, Width = 80, Height = 26 };
            grpTarget.Controls.Add(txtTargetDir);
            grpTarget.Controls.Add(btnBrowseTarget);
            tab.Controls.Add(grpTarget);
            btnBrowseTarget.Click += BtnBrowseTarget_Click;

            y += 80;

            // 开始按钮
            btnStartCopy = new Button { Text = "开始查找并复制", Left = 15, Top = y, Width = 140, Height = 32 };
            tab.Controls.Add(btnStartCopy);
            btnStartCopy.Click += BtnStartCopy_Click;

            y += 45;

            // 进度与日志
            pgbCopy = new ProgressBar { Left = 15, Top = y, Width = 900, Height = 18 };
            lblCopyStatus = new Label { Left = 15, Top = y + 25, Width = 300, Height = 20, Text = "就绪" };
            lstLog = new ListBox { Left = 15, Top = y + 55, Width = 900, Height = 200, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom, ScrollAlwaysVisible = true };
            tab.Controls.Add(pgbCopy);
            tab.Controls.Add(lblCopyStatus);
            tab.Controls.Add(lstLog);
        }

        // ---------- 浏览按钮 ----------
        private void BtnBrowseSource_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog { Description = "请选择包含表格的目标文件夹" })
            {
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    txtSourceFolder.Text = fbd.SelectedPath;
                    btnSummary.Enabled = true;
                    lblStatus.Text = "已选择源文件夹";
                }
            }
        }

        private void BtnBrowseSnSource_Click(object sender, EventArgs e)
        {
            var ext = rdoFromTxt.Checked ? "文本文件(*.txt)|*.txt|所有文件(*.*)|*.*" : "表格文件(*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx|所有文件(*.*)|*.*";
            using (var dlg = new OpenFileDialog { Filter = ext, Title = "选择 SN 来源文件" })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    txtSnSource.Text = dlg.FileName;
            }
        }

        private void BtnBrowseImageRoot_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog { Description = "请选择图片文件夹的根目录" })
            {
                if (fbd.ShowDialog(this) == DialogResult.OK)
                    txtImageRoot.Text = fbd.SelectedPath;
            }
        }

        private void BtnBrowseTarget_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog { Description = "请选择复制目标路径" })
            {
                if (fbd.ShowDialog(this) == DialogResult.OK)
                    txtTargetDir.Text = fbd.SelectedPath;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _appIconBitmap?.Dispose();
            base.OnFormClosing(e);
        }

        // 创建内嵌图标（同时显示在标题栏与任务栏）
        private Icon CreateAppIcon()
        {
            _appIconBitmap = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(_appIconBitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.Transparent);

                using (var bgBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    new Point(0, 0), new Point(0, 32),
                    Color.FromArgb(70, 130, 180), Color.FromArgb(50, 100, 150)))
                {
                    g.FillEllipse(bgBrush, 2, 2, 28, 28);
                }
                g.DrawEllipse(new Pen(Color.White, 1.5f), 2, 2, 27, 27);

                using (var font = new Font("Segoe UI", 10f, FontStyle.Bold))
                using (var brush = new SolidBrush(Color.White))
                {
                    var sf = new System.Drawing.StringFormat
                    {
                        Alignment = System.Drawing.StringAlignment.Center,
                        LineAlignment = System.Drawing.StringAlignment.Center
                    };
                    g.DrawString("E", font, brush, new RectangleF(0, 0, 16, 32), sf);
                }
            }
            return Icon.FromHandle(_appIconBitmap.GetHicon());
        }
    }
}
