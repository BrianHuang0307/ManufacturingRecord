using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System.Reflection;

namespace ManufacturingRecord
{
    public partial class Form1 : Form
    {
        // 改用 TableLayoutPanel
        private TableLayoutPanel _topTable;
        private TableLayoutPanel _filterTable;

        private Label _fromLabel;
        private DateTimePicker _fromDateTimePicker;
        private Label _toLabel;
        private DateTimePicker _toDateTimePicker;
        private Button _searchDataButton;
        private Button _exportExcelButton;
        private Label _announcementLabel;

        private Label _productLabel;
        private TextBox _productTextBox;
        private Label _productAbbreviationLabel;
        private TextBox _productAbbreviationTextBox;
        private Label _featureLabel;
        private TextBox _featureTextBox;
        private Label _processLabel;
        private TextBox _processTextBox;
        private Label _machineLabel;
        private TextBox _machineTextBox;

        private Button _selectDgvButton;
        private Button _calculateDailyCapacityButton;
        private Button _calculateAverageCapacityButton;
        private DataGridView dgv;

        private readonly IData _dataService;
        private readonly IExcelService _excelService;

        public Form1(IData dataService, IExcelService excelService)
        {
            InitializeComponent();
            _dataService = dataService;
            _excelService = excelService;

            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            // 設定初始大小，建議寬一點以容納所有欄位
            this.ClientSize = new Size(1400, 800);
            this.Text = "機器生產履歷";

            // --- 圖示設置 ---
            try
            {
                string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string iconPath = Path.Combine(appDirectory, "machine_production_history_icon.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading icon: {ex.Message}");
            }

            // ==========================================
            // 1. 初始化控制項 (不設定絕對位置與大小，交給 TableLayout)
            // ==========================================
            const int ControlHeight = 28; //稍微加高一點比較好看

            // --- 第一列控制項 ---
            _fromLabel = CreateLabel("開始日期");
            _fromDateTimePicker = CreateDateTimePicker(ControlHeight);
            _toLabel = CreateLabel("結束日期");
            _toDateTimePicker = CreateDateTimePicker(ControlHeight);
            _searchDataButton = CreateButton("查詢資料", ControlHeight);
            _exportExcelButton = CreateButton("匯出Excel", ControlHeight);
            _announcementLabel = new Label
            {
                Text = "※ 僅包含結案一般工單，計算產能前請先點選「選取資料」",
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                ForeColor = Color.Red,
                Anchor = AnchorStyles.Left | AnchorStyles.Right // 讓它跟著拉伸
            };

            // --- 第二列控制項 ---
            _productLabel = CreateLabel("生產料件");
            _productTextBox = CreateTextBox(ControlHeight);
            _productAbbreviationLabel = CreateLabel("品名簡稱");
            _productAbbreviationTextBox = CreateTextBox(ControlHeight);
            _featureLabel = CreateLabel("特性編碼");
            _featureTextBox = CreateTextBox(ControlHeight);
            _processLabel = CreateLabel("工序");
            _processTextBox = CreateTextBox(ControlHeight);
            _machineLabel = CreateLabel("機台編號");
            _machineTextBox = CreateTextBox(ControlHeight);

            _selectDgvButton = CreateButton("選取資料", ControlHeight);

            _calculateDailyCapacityButton = CreateButton("工單日產能計算", ControlHeight);
            _calculateDailyCapacityButton.BackColor = Color.MistyRose;
            _calculateDailyCapacityButton.FlatStyle = FlatStyle.Flat;

            _calculateAverageCapacityButton = CreateButton("平均工單日產能計算", ControlHeight);
            _calculateAverageCapacityButton.BackColor = Color.Honeydew;
            _calculateAverageCapacityButton.FlatStyle = FlatStyle.Flat;

            // ==========================================
            // 2. 建立佈局 (TableLayoutPanel)
            // ==========================================

            // --- 頂部佈局 (第一列) ---
            // 欄位規劃: [Label][Picker][Label][Picker][Btn][Btn][Announcement(剩下的空間)]
            _topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 1,
                ColumnCount = 7,
                Padding = new Padding(5),
                Height = 40,
                BackColor = Color.AliceBlue
            };

            // 設定欄位比例 (關鍵!)
            // Label 用 AutoSize，Input 用 Percent (讓它們負責縮放)，Button 用 AutoSize
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 開始標籤
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F)); // 開始時間 (15%)
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 結束標籤
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F)); // 結束時間 (15%)
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 查詢鈕
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 匯出鈕
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F)); // 公告 (佔據剩下所有空間)

            // 加入控制項到 _topTable
            _topTable.Controls.Add(_fromLabel, 0, 0);
            _topTable.Controls.Add(_fromDateTimePicker, 1, 0);
            _topTable.Controls.Add(_toLabel, 2, 0);
            _topTable.Controls.Add(_toDateTimePicker, 3, 0);
            _topTable.Controls.Add(_searchDataButton, 4, 0);
            _topTable.Controls.Add(_exportExcelButton, 5, 0);
            _topTable.Controls.Add(_announcementLabel, 6, 0);


            // --- 過濾器佈局 (第二列) ---
            // 欄位規劃: [Lbl][Txt] * 5組 + [SelectBtn] + [DailyBtn] + [AvgBtn]
            // 共 11 個欄位
            _filterTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 1,
                ColumnCount = 13,
                Padding = new Padding(5),
                Height = 40,
                BackColor = Color.AliceBlue
            };

            // 設定欄位比例
            // 文字框設為 20% 讓它們均分寬度，按鈕設為 AutoSize 或固定比例
            for (int i = 0; i < 5; i++)
            {
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 標籤 (Auto)
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F)); // 輸入框 (20%)
            }
            // 3個按鈕
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // 加入控制項到 _filterTable
            _filterTable.Controls.Add(_productLabel, 0, 0);
            _filterTable.Controls.Add(_productTextBox, 1, 0);
            _filterTable.Controls.Add(_productAbbreviationLabel, 2, 0);
            _filterTable.Controls.Add(_productAbbreviationTextBox, 3, 0);
            _filterTable.Controls.Add(_featureLabel, 4, 0);
            _filterTable.Controls.Add(_featureTextBox, 5, 0);
            _filterTable.Controls.Add(_processLabel, 6, 0);
            _filterTable.Controls.Add(_processTextBox, 7, 0);
            _filterTable.Controls.Add(_machineLabel, 8, 0);
            _filterTable.Controls.Add(_machineTextBox, 9, 0);

            _filterTable.Controls.Add(_selectDgvButton, 10, 0);
            _filterTable.Controls.Add(_calculateDailyCapacityButton, 11, 0);
            _filterTable.Controls.Add(_calculateAverageCapacityButton, 12, 0);


            // --- DataGridView ---
            dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackgroundColor = Color.AliceBlue,
                BorderStyle = BorderStyle.Fixed3D,
                Margin = new Padding(10),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoGenerateColumns = true,
                AllowUserToOrderColumns = true
            };

            // ==========================================
            // 3. 加入 Form
            // ==========================================
            // 注意順序：Dock 是 "堆疊" 的概念。
            // 先加的 Dock=Top 會在最上面，後加的 Dock=Top 會接在下面。
            // 但如果全部一起加，通常反過來思考：
            // 我們希望 dgv 填滿中間，TopTable 在最上，FilterTable 在中間。

            // 這裡使用最直覺的方式：依照順序加入 Controls 集合，
            // 但因為 Dock 屬性的特性，最後加入的 Dock.Fill 會佔據剩餘空間。

            this.Controls.Add(dgv);          // Fill (最底層)
            this.Controls.Add(_filterTable); // Top (第二層，卡在 TopTable 下面)
            this.Controls.Add(_topTable);    // Top (最上層)

            EnableDoubleBuffered(dgv);
            AddEventHandlers();
        }

        // --- 簡化後的輔助方法 (不再需要傳入 width/margin/position) ---

        private Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Right, // 靠右對齊 (靠近輸入框)
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(5, 0, 0, 0) // 左邊留點空隙
            };
        }

        private TextBox CreateTextBox(int height)
        {
            return new TextBox
            {
                Height = height,
                Dock = DockStyle.Fill, // 填滿格子
                Margin = new Padding(3)
            };
        }

        private DateTimePicker CreateDateTimePicker(int height)
        {
            return new DateTimePicker
            {
                Height = height,
                Dock = DockStyle.Fill, // 填滿格子
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd",
                Margin = new Padding(3)
            };
        }

        private Button CreateButton(string text, int height)
        {
            return new Button
            {
                Text = text,
                Height = height,
                AutoSize = true, // 讓按鈕寬度隨文字自動調整
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(10, 0, 10, 0), // 增加內部文字間距
                Margin = new Padding(3),
                Cursor = Cursors.Hand
            };
        }

        // 避免畫面在resize時因為資料量過多，造成卡頓情形
        private void EnableDoubleBuffered(DataGridView dgv)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi?.SetValue(dgv, true, null);
        }
    }
}