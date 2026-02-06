using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ManufacturingRecord
{
    public partial class ProductErrorCodeView : UserControl
    {
        // UI 控制項 (原本在 Form1 的都搬來這裡)
        private TableLayoutPanel _topTable = null!;
        private TableLayoutPanel _filterTable = null!;
        private Label _announcementLabel = null!;
        private Button _searchDataButton = null!; 
        private Button _exportExcelButton = null!;

        private Label _productLabel = null!;
        private Label _featureLabel = null!;
        private Label _processLabel = null!;
        // private Label _productAbbreviationLabel = null!;

        private TextBox _productTextBox = null!;
        private TextBox _featureTextBox = null!;
        private TextBox _processTextBox = null!;
        // private TextBox _productAbbreviationTextBox = null!;

        private Button _selectDgvButton = null!;
        private Button _calculateButton = null!;
        private DataGridView dgv = null!;

        // Service 變數
        private IData _dataService = null!;
        private IExcelService _excelService = null!;
        public ProductErrorCodeView(IData dataService, IExcelService excelService)
        {
            InitializeComponent();
            _dataService = dataService;
            _excelService = excelService;

            // 初始化介面
            InitializeCustomUI();

            AddEventHandlers();
        }

        public ProductErrorCodeView()
        {
            InitializeComponent();
        }

        private void InitializeCustomUI()
        {
            this.DoubleBuffered = true;
            this.BackColor = Color.White; // UserControl 預設背景色

            const int ControlHeight = 28;

            // --- 建立控制項 (邏輯同原本 Form1) ---
            _searchDataButton = CreateButton("載入資料", ControlHeight);
            _exportExcelButton = CreateButton("匯出Excel", ControlHeight);

            _announcementLabel = new Label
            {
                Text = "※ 僅包含結案一般工單，可用，或、或;進行多重篩選",
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                ForeColor = Color.Red,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };

            _productLabel = CreateLabel("生產料件");
            _productTextBox = CreateTextBox(ControlHeight);

            /*
            _productAbbreviationLabel = CreateLabel("品名簡稱");
            _productAbbreviationTextBox = CreateTextBox(ControlHeight);
            */

            _featureLabel = CreateLabel("特性編碼");
            _featureTextBox = CreateTextBox(ControlHeight);
            _processLabel = CreateLabel("工序");
            _processTextBox = CreateTextBox(ControlHeight);
 
            _selectDgvButton = CreateButton("選取資料", ControlHeight);
            _calculateButton = CreateButton("計算", ControlHeight);

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
                ColumnCount = 5,
                Padding = new Padding(5),
                Height = 40,
                BackColor = Color.AliceBlue
            };
            // ... (設定 _topTable ColumnStyles) ...
            for (int i = 0; i < 3; i++)
            {
                _topTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            }

            _topTable.Controls.Add(_searchDataButton, 0, 0);
            _topTable.Controls.Add(_exportExcelButton, 1, 0);
            _topTable.Controls.Add(_announcementLabel, 2, 0);

            _filterTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 1,
                ColumnCount = 8,
                Padding = new Padding(5),
                Height = 40,
                BackColor = Color.AliceBlue
            };

            // ... (設定 _filterTable ColumnStyles) ...
            for (int i = 0; i < 3; i++)
            {
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            }

            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _filterTable.Controls.Add(_productLabel, 0, 0);
            _filterTable.Controls.Add(_productTextBox, 1, 0);
            _filterTable.Controls.Add(_featureLabel, 2, 0);
            _filterTable.Controls.Add(_featureTextBox, 3, 0);
            _filterTable.Controls.Add(_processLabel, 4, 0);
            _filterTable.Controls.Add(_processTextBox, 5, 0);
            _filterTable.Controls.Add(_selectDgvButton, 6, 0);
            _filterTable.Controls.Add(_calculateButton, 7, 0);
            /*
            _filterTable.Controls.Add(_productAbbreviationLabel, 2, 0);
            _filterTable.Controls.Add(_productAbbreviationTextBox, 3, 0);
            _filterTable.Controls.Add(_featureLabel, 4, 0);
            _filterTable.Controls.Add(_featureTextBox, 5, 0);
            _filterTable.Controls.Add(_processLabel, 6, 0);
            _filterTable.Controls.Add(_processTextBox, 7, 0);
            _filterTable.Controls.Add(_selectDgvButton, 8, 0);
            */

            // 加入 UserControl
            this.Controls.Add(dgv);
            this.Controls.Add(_filterTable);
            this.Controls.Add(_topTable);
        }

        // ... (保留 CreateButton, CreateLabel 等輔助方法) ...
        private Button CreateButton(string text, int height) { return new Button { Text = text, Height = height, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10, 0, 10, 0), Margin = new Padding(3), Cursor = Cursors.Hand }; }
        private Label CreateLabel(string text) { return new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Right, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5, 0, 0, 0) }; }
        private TextBox CreateTextBox(int height) { return new TextBox { Height = height, Dock = DockStyle.Fill, Margin = new Padding(3) }; }
        private void EnableDoubleBuffered(DataGridView dgv) { Type dgvType = dgv.GetType(); PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic); pi?.SetValue(dgv, true, null); }
    }
}
