using System.Drawing;
using System.Windows.Forms;
using System;
using System.IO;
using System.Reflection;

namespace ManufacturingRecord
{
    public partial class Form1 : Form
    {
        private FlowLayoutPanel _flowLayoutPanel;
        private Panel _fromDateGroupPanel;
        private Panel _toDateGroupPanel;
        //private Panel _productGroupPanel;
        //private Panel _featureGroupPanel;
        //private Panel _processGroupPanel;
        //private Label _productLabel;
        //private TextBox _productTextBox;
        //private Label _featureLabel;
        //private TextBox _featureTextBox;
        //private Label _processLabel;
        //private TextBox _processTextBox;
        private Label _fromLabel;
        private DateTimePicker _fromDateTimePicker;
        private Label _toLabel;
        private DateTimePicker _toDateTimePicker;
        private Button _searchButton;
        private Button _exportExcelButton;
        private DataGridView dgv;

        public Form1()
        {
            // 假設此方法定義在 Form1.Designer.cs 中
            InitializeComponent();

            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(600, 800);
            this.Text = "機器生產履歷";

            // --- 輔助常數設定 ---
            const int LabelWidth = 55;
            const int ControlHeight = 25;
            const int ControlWidth = 100;
            const int LabelRightMargin = 5;
            const int ButtonWidth = 80;

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

            _flowLayoutPanel = new FlowLayoutPanel
            {
                AutoScroll = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(5, 10, 5, 10),
                Dock = DockStyle.Top,
                BackColor = Color.AliceBlue,
                BorderStyle = BorderStyle.Fixed3D,
            };

            /*
            // --- 產品輸入組件 ---
            _productLabel = CreateLabel(LabelWidth, ControlHeight, "生產料件", LabelRightMargin);
            _productTextBox = CreateTextBox(ControlWidth, ControlHeight, 0);
            _productGroupPanel = CreateGroupPanel(_productLabel, _productTextBox, ControlHeight);

            // --- 特性編碼輸入組件 ---
            _featureLabel = CreateLabel(LabelWidth, ControlHeight, "特性編碼", LabelRightMargin);
            _featureTextBox = CreateTextBox(ControlWidth, ControlHeight, 0);
            _featureGroupPanel = CreateGroupPanel(_featureLabel, _featureTextBox, ControlHeight);

            // --- 工序輸入組件 ---
            _processLabel = CreateLabel(LabelWidth, ControlHeight, "工序", LabelRightMargin);
            _processTextBox = CreateTextBox(ControlWidth, ControlHeight, 0);
            _processGroupPanel = CreateGroupPanel(_processLabel, _processTextBox, ControlHeight);
            */

            // --- 開始日期組件 ---
            _fromLabel = CreateLabel(LabelWidth, ControlHeight, "開始", LabelRightMargin);
            _fromDateTimePicker = CreateDateTimePicker(ControlWidth, ControlHeight, 0);
            _fromDateGroupPanel = CreateGroupPanel(_fromLabel, _fromDateTimePicker, ControlHeight);

            // --- 結束日期組件 ---
            _toLabel = CreateLabel(LabelWidth, ControlHeight, "結束", LabelRightMargin);
            _toDateTimePicker = CreateDateTimePicker(ControlWidth, ControlHeight, 0);
            _toDateGroupPanel = CreateGroupPanel(_toLabel, _toDateTimePicker, ControlHeight);

            // --- 查詢按鈕 ---
            _searchButton = new Button
            {
                Text = "查詢",
                Width = ButtonWidth,
                Height = ControlHeight,
                Margin = new Padding(5, 10, 5, 10),
            };

            _exportExcelButton = new Button
            {
                Text = "匯出Excel",
                Width = ButtonWidth,
                Height = ControlHeight,
                Margin = new Padding(15, 10, 15, 10),
            };

            // 1. 將查詢條件 FlowLayoutPanel 加入 Form
            Controls.Add(_flowLayoutPanel);
            //_flowLayoutPanel.Controls.Add(_productGroupPanel);
            //_flowLayoutPanel.Controls.Add(_featureGroupPanel);
            //_flowLayoutPanel.Controls.Add(_processGroupPanel);
            _flowLayoutPanel.Controls.Add(_fromDateGroupPanel);
            _flowLayoutPanel.Controls.Add(_toDateGroupPanel);
            _flowLayoutPanel.Controls.Add(_searchButton);
            _flowLayoutPanel.Controls.Add(_exportExcelButton);

            // 2. 新增 DataGridView
            dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                BackgroundColor = Color.AliceBlue,
                BorderStyle = BorderStyle.Fixed3D,
                Margin = new Padding(10, 10, 10, 10),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoGenerateColumns = true
            };

            // 3. 將 dgv 加入 Form
            Controls.Add(dgv);
            dgv.BringToFront();
            EnableDoubleBuffered(dgv);

            AddEventHandlers();
        }

        /// <summary>
        /// 輔助方法：創建一個 Panel 來包含 Label 和 Control，並進行手動定位。
        /// **修正：明確指定 control 參數的類型為 System.Windows.Forms.Control**
        /// </summary>
        private Panel CreateGroupPanel(Label label, System.Windows.Forms.Control control, int controlHeight)
        {
            Panel panel = new Panel
            {
                Margin = new Padding(5, 10, 5, 10),
                Height = controlHeight,
                Width = label.Width + control.Width + label.Margin.Right + 5,
            };

            panel.Controls.Add(label);
            panel.Controls.Add(control);

            label.Top = 0;
            control.Left = label.Right + label.Margin.Right;
            control.Top = 0;

            return panel;
        }

        /// <summary>
        /// 輔助方法：統一設置 Label 的屬性。
        /// </summary>
        private Label CreateLabel(int width, int height, string labelName, int rightMargin)
        {
            Label label = new Label
            {
                Width = width,
                Height = height,
                Text = labelName,
                Margin = new Padding(0, 0, rightMargin, 0),
                TextAlign = ContentAlignment.MiddleRight
            };

            return label;
        }

        /*
        private TextBox CreateTextBox(int width, int height, int margin)
        {
            TextBox textBox = new TextBox
            {
                Width = width,
                Height = height,
                Margin = new Padding(margin)
            };

            return textBox;
        }
        */

        private DateTimePicker CreateDateTimePicker(int width, int height, int margin)
        {
            DateTimePicker dateTimePicker = new DateTimePicker
            {
                Width = width,
                Height = height,
                Margin = new Padding(margin),
                CustomFormat = "yyyy-MM-dd"
            };

            return dateTimePicker;
        }

        // 解決視窗resize時卡頓問題
        private void EnableDoubleBuffered(DataGridView dgv)
        {
            // 利用反射取得 DataGridView 的 DoubleBuffered 屬性
            Type dgvType = dgv.GetType();
            PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);

            // 強制設為 true
            pi?.SetValue(dgv, true, null);
        }
    }
}