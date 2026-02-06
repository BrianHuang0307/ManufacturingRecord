using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System.Reflection;

namespace ManufacturingRecord
{
    // 注意：要是 partial class
    public partial class MachineRecordView : UserControl
    {
        // UI 控制項 (原本在 Form1 的都搬來這裡)
        private TableLayoutPanel _topTable = null!;
        private TableLayoutPanel _filterTable = null!;
        // ... (省略部分變數宣告，跟你原本一樣，只是權限變成 private) ...
        private Label _fromLabel = null!;
        private Label _toLabel = null!; 
        private Label _announcementLabel = null!;
        private DateTimePicker _fromDateTimePicker = null!;
        private DateTimePicker _toDateTimePicker = null!;
        private Button _searchDataButton = null!;
        private Button _exportExcelButton = null!;

        private Label _productLabel = null!;
        private Label _productAbbreviationLabel = null!;
        private Label _groundLabel = null!;
        private Label _ceilLabel = null!;
        private Label _featureLabel = null!;
        private Label _processLabel = null!;
        private Label _machineLabel = null!;

        private TextBox _productTextBox = null!;
        private TextBox _productAbbreviationTextBox = null!;
        private TextBox _groundTextBox = null!;
        private TextBox _ceilTextBox = null!;
        private TextBox _featureTextBox = null!;
        private TextBox _processTextBox = null!;
        private TextBox _machineTextBox = null!;

        private Button _selectDgvButton = null!;
        private Button _calculateDailyCapacityButton = null!;
        private Button _calculateAverageCapacityButton = null!;
        private DataGridView dgv = null!;

        // Service 變數
        private IData _dataService = null!;
        private IExcelService _excelService = null!;

        // 建構函式：接收 Service
        public MachineRecordView(IData dataService, IExcelService excelService)
        {
            InitializeComponent(); // 手寫UI，保留習慣
            _dataService = dataService;
            _excelService = excelService;

            // 初始化介面
            InitializeCustomUI();

            // 綁定事件 (這個方法會定義在 MachineRecord.Event.cs)
            AddEventHandlers();
        }

        // 為了避免設計檢視器報錯，保留一個無參數建構函式 (可選)
        public MachineRecordView()
        {
            InitializeComponent();
        }

        private void InitializeCustomUI()
        {
            this.DoubleBuffered = true;
            this.BackColor = Color.White; // UserControl 預設背景色

            const int ControlHeight = 28;

            // --- 建立控制項 (邏輯同原本 Form1) ---
            _fromLabel = CreateLabel("開始日期");
            _fromDateTimePicker = CreateDateTimePicker(ControlHeight);
            _toLabel = CreateLabel("結束日期");
            _toDateTimePicker = CreateDateTimePicker(ControlHeight);
            _searchDataButton = CreateButton("查詢資料", ControlHeight);
            _exportExcelButton = CreateButton("匯出Excel", ControlHeight);

            _announcementLabel = new Label
            {
                Text = "※ 僅包含結案一般工單，計算產能前請先點選「選取資料」，可用，或、或;進行多重篩選",
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                ForeColor = Color.Red,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };

            _groundLabel = CreateLabel("機器工時下限(分鐘)");
            _groundTextBox = CreateTextBox(ControlHeight);
            _ceilLabel = CreateLabel("機器工時上限(分鐘)");
            _ceilTextBox = CreateTextBox(ControlHeight);

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
            EnableDoubleBuffered(dgv);

            // --- Layout ---
            _topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 1,
                ColumnCount = 11,
                Padding = new Padding(5),
                Height = 40,
                BackColor = Color.AliceBlue
            };
            // ... (設定 _topTable ColumnStyles) ...
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15F));

            _topTable.Controls.Add(_fromLabel, 0, 0);
            _topTable.Controls.Add(_fromDateTimePicker, 1, 0);
            _topTable.Controls.Add(_toLabel, 2, 0);
            _topTable.Controls.Add(_toDateTimePicker, 3, 0);
            _topTable.Controls.Add(_searchDataButton, 4, 0);
            _topTable.Controls.Add(_exportExcelButton, 5, 0);
            _topTable.Controls.Add(_announcementLabel, 6, 0);
            _topTable.Controls.Add(_groundLabel, 7, 0);
            _topTable.Controls.Add(_groundTextBox, 8, 0);
            _topTable.Controls.Add(_ceilLabel, 9, 0);
            _topTable.Controls.Add(_ceilTextBox, 10, 0);

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
            // ... (設定 _filterTable ColumnStyles) ...
            for (int i = 0; i < 5; i++)
            {
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            }
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

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

            // 加入 UserControl
            this.Controls.Add(dgv);
            this.Controls.Add(_filterTable);
            this.Controls.Add(_topTable);
        }

        // ... (保留 CreateButton, CreateLabel 等輔助方法) ...
        private Button CreateButton(string text, int height) { return new Button { Text = text, Height = height, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10, 0, 10, 0), Margin = new Padding(3), Cursor = Cursors.Hand }; }
        private Label CreateLabel(string text) { return new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Right, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5, 0, 0, 0) }; }
        private TextBox CreateTextBox(int height) { return new TextBox { Height = height, Dock = DockStyle.Fill, Margin = new Padding(3) }; }
        private DateTimePicker CreateDateTimePicker(int height) { return new DateTimePicker { Height = height, Dock = DockStyle.Fill, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Margin = new Padding(3) }; }
        private void EnableDoubleBuffered(DataGridView dgv) { Type dgvType = dgv.GetType(); PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic); pi?.SetValue(dgv, true, null); }
    }
}