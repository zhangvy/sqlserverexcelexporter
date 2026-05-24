using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SqlServerExcelExporter.Core;

namespace SqlServerExcelExporter.App
{
    public sealed class MainForm : Form
    {
        private readonly TextBox _tableNameTextBox = new TextBox();
        private readonly Button _loadColumnsButton = new Button();
        private readonly DataGridView _columnsGrid = new DataGridView();
        private readonly ComboBox _dateColumnComboBox = new ComboBox();
        private readonly DateTimePicker _startDatePicker = new DateTimePicker();
        private readonly DateTimePicker _endDatePicker = new DateTimePicker();
        private readonly ComboBox _groupModeComboBox = new ComboBox();
        private readonly TextBox _outputDirectoryTextBox = new TextBox();
        private readonly Button _browseButton = new Button();
        private readonly Button _exportButton = new Button();
        private readonly ProgressBar _progressBar = new ProgressBar();
        private readonly TextBox _logTextBox = new TextBox();

        private IList<ColumnInfo> _columns = new List<ColumnInfo>();

        public MainForm()
        {
            Text = "SQL Server 导出 Excel";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(980, 680);
            Size = new Size(1080, 760);

            BuildLayout();
            LoadDefaults();
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 116));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            Controls.Add(root);

            root.Controls.Add(BuildTopPanel(), 0, 0);
            root.Controls.Add(BuildColumnsGrid(), 0, 1);
            root.Controls.Add(BuildExportPanel(), 0, 2);
            root.Controls.Add(BuildLogPanel(), 0, 3);
        }

        private Control BuildTopPanel()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            panel.Controls.Add(CreateLabel("表名"), 0, 0);
            panel.Controls.Add(CreateHintLabel("格式：dbo.TableName；只填 TableName 时默认 dbo"), 1, 0);

            _tableNameTextBox.Dock = DockStyle.Fill;
            _tableNameTextBox.Margin = new Padding(0, 4, 8, 4);
            _tableNameTextBox.Text = "dbo.TableName";
            panel.Controls.Add(_tableNameTextBox, 1, 1);

            _loadColumnsButton.Text = "查询表结构";
            _loadColumnsButton.Dock = DockStyle.Fill;
            _loadColumnsButton.Margin = new Padding(0, 4, 0, 4);
            _loadColumnsButton.Click += LoadColumnsButton_Click;
            panel.Controls.Add(_loadColumnsButton, 2, 1);

            return panel;
        }

        private Control BuildColumnsGrid()
        {
            _columnsGrid.Dock = DockStyle.Fill;
            _columnsGrid.AllowUserToAddRows = false;
            _columnsGrid.AllowUserToDeleteRows = false;
            _columnsGrid.AllowUserToResizeRows = false;
            _columnsGrid.AutoGenerateColumns = false;
            _columnsGrid.ReadOnly = true;
            _columnsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _columnsGrid.MultiSelect = false;
            _columnsGrid.RowHeadersVisible = false;
            _columnsGrid.BackgroundColor = SystemColors.Window;

            _columnsGrid.Columns.Add(CreateTextColumn("Ordinal", "顺序", 60));
            _columnsGrid.Columns.Add(CreateTextColumn("Name", "字段名", 220));
            _columnsGrid.Columns.Add(CreateTextColumn("DataType", "类型", 140));
            _columnsGrid.Columns.Add(CreateTextColumn("IsNullable", "可空", 70));
            _columnsGrid.Columns.Add(CreateTextColumn("MaxLength", "长度", 80));
            _columnsGrid.Columns.Add(CreateTextColumn("NumericPrecision", "精度", 80));
            _columnsGrid.Columns.Add(CreateTextColumn("NumericScale", "小数位", 80));
            _columnsGrid.Columns.Add(CreateTextColumn("IsDateColumn", "日期列", 80));

            return _columnsGrid;
        }

        private Control BuildExportPanel()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 3,
                Padding = new Padding(0, 8, 0, 0)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

            _dateColumnComboBox.Dock = DockStyle.Fill;
            _dateColumnComboBox.DropDownStyle = ComboBoxStyle.DropDownList;

            _startDatePicker.Dock = DockStyle.Fill;
            _startDatePicker.Format = DateTimePickerFormat.Custom;
            _startDatePicker.CustomFormat = "yyyy-MM-dd";

            _endDatePicker.Dock = DockStyle.Fill;
            _endDatePicker.Format = DateTimePickerFormat.Custom;
            _endDatePicker.CustomFormat = "yyyy-MM-dd";

            _groupModeComboBox.Dock = DockStyle.Fill;
            _groupModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _groupModeComboBox.Items.Add(new GroupModeItem("不分组", ExportGroupMode.None));
            _groupModeComboBox.Items.Add(new GroupModeItem("按周", ExportGroupMode.Week));
            _groupModeComboBox.Items.Add(new GroupModeItem("按月", ExportGroupMode.Month));
            _groupModeComboBox.SelectedIndex = 0;

            _outputDirectoryTextBox.Dock = DockStyle.Fill;

            _browseButton.Text = "浏览";
            _browseButton.Dock = DockStyle.Fill;
            _browseButton.Click += BrowseButton_Click;

            _exportButton.Text = "开始导出";
            _exportButton.Dock = DockStyle.Fill;
            _exportButton.Click += ExportButton_Click;

            _progressBar.Dock = DockStyle.Fill;
            _progressBar.Style = ProgressBarStyle.Continuous;

            panel.Controls.Add(CreateLabel("日期列"), 0, 0);
            panel.Controls.Add(_dateColumnComboBox, 1, 0);
            panel.Controls.Add(CreateLabel("起始日期"), 2, 0);
            panel.Controls.Add(_startDatePicker, 3, 0);
            panel.Controls.Add(CreateLabel("结束日期"), 4, 0);
            panel.Controls.Add(_endDatePicker, 5, 0);

            panel.Controls.Add(CreateLabel("导出方式"), 0, 1);
            panel.Controls.Add(_groupModeComboBox, 1, 1);
            panel.Controls.Add(CreateLabel("输出目录"), 2, 1);
            panel.SetColumnSpan(_outputDirectoryTextBox, 2);
            panel.Controls.Add(_outputDirectoryTextBox, 3, 1);
            panel.Controls.Add(_browseButton, 5, 1);

            panel.Controls.Add(_exportButton, 0, 2);
            panel.SetColumnSpan(_progressBar, 5);
            panel.Controls.Add(_progressBar, 1, 2);

            return panel;
        }

        private Control BuildLogPanel()
        {
            _logTextBox.Dock = DockStyle.Fill;
            _logTextBox.Multiline = true;
            _logTextBox.ReadOnly = true;
            _logTextBox.ScrollBars = ScrollBars.Vertical;
            _logTextBox.Font = new Font("Consolas", 9F);
            return _logTextBox;
        }

        private void LoadDefaults()
        {
            var defaultDirectory = ConfigurationManager.AppSettings["DefaultOutputDirectory"];
            if (!string.IsNullOrWhiteSpace(defaultDirectory))
            {
                _outputDirectoryTextBox.Text = defaultDirectory;
            }

            _startDatePicker.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            _endDatePicker.Value = DateTime.Today;
        }

        private async void LoadColumnsButton_Click(object sender, EventArgs e)
        {
            await RunBusyAsync(async () =>
            {
                var tableName = TableName.Parse(_tableNameTextBox.Text);
                Log("正在查询表结构：" + tableName.DisplayName);

                var repository = CreateRepository();
                var columns = await Task.Run(() => repository.GetColumns(tableName));

                _columns = columns;
                _columnsGrid.DataSource = new BindingList<ColumnInfo>(columns.ToList());

                var dateColumns = columns.Where(c => c.IsDateColumn).ToList();
                _dateColumnComboBox.DataSource = dateColumns;
                _dateColumnComboBox.DisplayMember = "Name";
                _dateColumnComboBox.ValueMember = "Name";

                if (dateColumns.Count > 0)
                {
                    _dateColumnComboBox.SelectedIndex = 0;
                }

                Log("查询完成，共 " + columns.Count + " 个字段，日期列 " + dateColumns.Count + " 个。");
            });
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择 Excel 输出目录";
                dialog.SelectedPath = _outputDirectoryTextBox.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _outputDirectoryTextBox.Text = dialog.SelectedPath;
                }
            }
        }

        private async void ExportButton_Click(object sender, EventArgs e)
        {
            await RunBusyAsync(async () =>
            {
                var tableName = TableName.Parse(_tableNameTextBox.Text);
                var dateColumn = _dateColumnComboBox.SelectedItem as ColumnInfo;
                if (dateColumn == null)
                {
                    throw new InvalidOperationException("请先查询表结构并选择日期列。");
                }

                var mode = ((GroupModeItem)_groupModeComboBox.SelectedItem).Mode;
                var startDate = _startDatePicker.Value.Date;
                var endDate = _endDatePicker.Value.Date;
                var outputDirectory = _outputDirectoryTextBox.Text.Trim();

                _progressBar.Value = 0;
                Log("开始导出：" + tableName.DisplayName);
                Log("日期范围：" + startDate.ToString("yyyy-MM-dd") + " 到 " + endDate.ToString("yyyy-MM-dd") + "（包含结束日期）");

                var repository = CreateRepository();
                var exportService = new ExportService(repository, new XlsxWriter());
                var progress = new Progress<ExportProgress>(UpdateExportProgress);

                var results = await Task.Run(() => exportService.Export(
                    tableName,
                    dateColumn.Name,
                    dateColumn.DateColumnType,
                    startDate,
                    endDate,
                    mode,
                    outputDirectory,
                    progress));

                foreach (var result in results)
                {
                    Log("完成：" + result.Period.Label + "，" + result.RowCount + " 行，" + result.FilePath);
                }

                MessageBox.Show(this, "导出完成，共生成 " + results.Count + " 个 Excel 文件。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }

        private SqlServerRepository CreateRepository()
        {
            var setting = ConfigurationManager.ConnectionStrings["DefaultConnection"];
            if (setting == null || string.IsNullOrWhiteSpace(setting.ConnectionString))
            {
                throw new InvalidOperationException("请在 App.config 的 connectionStrings 中配置 DefaultConnection。");
            }

            return new SqlServerRepository(setting.ConnectionString);
        }

        private async Task RunBusyAsync(Func<Task> action)
        {
            SetBusy(true);
            try
            {
                await action();
            }
            catch (Exception ex)
            {
                Log("错误：" + ex.Message);
                MessageBox.Show(this, ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void SetBusy(bool busy)
        {
            _loadColumnsButton.Enabled = !busy;
            _exportButton.Enabled = !busy;
            _browseButton.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private void UpdateExportProgress(ExportProgress progress)
        {
            if (progress.TotalPeriods > 0)
            {
                _progressBar.Maximum = progress.TotalPeriods;
                _progressBar.Value = Math.Min(progress.CurrentPeriod, progress.TotalPeriods);
            }

            Log(progress.Message);
        }

        private void Log(string message)
        {
            _logTextBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };
        }

        private static Label CreateHintLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                AutoSize = false,
                ForeColor = SystemColors.GrayText
            };
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(string propertyName, string headerText, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                DataPropertyName = propertyName,
                HeaderText = headerText,
                Width = width,
                SortMode = DataGridViewColumnSortMode.Automatic
            };
        }

        private sealed class GroupModeItem
        {
            public GroupModeItem(string label, ExportGroupMode mode)
            {
                Label = label;
                Mode = mode;
            }

            public string Label { get; private set; }

            public ExportGroupMode Mode { get; private set; }

            public override string ToString()
            {
                return Label;
            }
        }
    }
}
