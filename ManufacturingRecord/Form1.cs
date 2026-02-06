using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System.Reflection;

namespace ManufacturingRecord
{
    public partial class Form1 : Form
    {
        // 導航列與按鈕
        private FlowLayoutPanel _navPanel = null!;
        private Button _navProductionPlannerButton = null!;
        private Button _navQualityEngineerButton = null!;

        // 內容容器 (用來放 UserControl)
        private Panel _contentPanel = null!;

        // 預先建立好的 Views (這樣切換回來時，查詢結果還會在，不會被清空)
        private MachineRecordView _machineRecordView = null!;
        private ProductErrorCodeView _productErrorCodeView = null!;

        private readonly IData _dataService;
        private readonly IExcelService _excelService;

        public Form1(IData dataService, IExcelService excelService)
        {
            InitializeComponent();
            _dataService = dataService;
            _excelService = excelService;

            // 基本視窗設定
            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(1400, 800);
            this.Text = "廠內生產管理"; // 改個通用的標題

            LoadIcon();

            // 1. 初始化導航列
            InitNavigation();

            // 2. 初始化內容容器
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            // 3. 初始化各個分頁 (將 Service 傳入)
            _machineRecordView = new MachineRecordView(_dataService, _excelService) { Dock = DockStyle.Fill };
            _productErrorCodeView = new ProductErrorCodeView(_dataService, _excelService) { Dock = DockStyle.Fill };

            // 4. 加入控制項 (注意順序：Nav 在上，Content 填滿剩餘)
            this.Controls.Add(_contentPanel);
            this.Controls.Add(_navPanel);

            // 5. 預設顯示第一個分頁
            SwitchTab(_navProductionPlannerButton);
        }

        private void InitNavigation()
        {
            _navProductionPlannerButton = CreateNavButton("生管日產能計算", true);
            _navQualityEngineerButton = CreateNavButton("品管錯誤代碼統計", false);

            _navProductionPlannerButton.Click += (s, e) => SwitchTab(_navProductionPlannerButton);
            _navQualityEngineerButton.Click += (s, e) => SwitchTab(_navQualityEngineerButton);

            _navPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 35,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(5, 2, 0, 0)
            };

            _navPanel.Controls.Add(_navProductionPlannerButton);
            _navPanel.Controls.Add(_navQualityEngineerButton);
        }

        private void SwitchTab(Button clickedBtn)
        {
            // 1. UI 樣式切換
            foreach (Control c in _navPanel.Controls)
            {
                if (c is Button btn)
                {
                    bool isSelected = (btn == clickedBtn);
                    btn.BackColor = isSelected ? Color.CornflowerBlue : Color.LightGray;
                    btn.ForeColor = isSelected ? Color.White : Color.Black;
                }
            }

            // 2. 內容切換
            _contentPanel.Controls.Clear(); // 清空目前顯示的

            if (clickedBtn == _navProductionPlannerButton)
            {
                _contentPanel.Controls.Add(_machineRecordView);
            }
            else if (clickedBtn == _navQualityEngineerButton)
            {
                _contentPanel.Controls.Add(_productErrorCodeView);
            }
        }

        private Button CreateNavButton(string text, bool isActive)
        {
            return new Button
            {
                Text = text,
                Width = 140, // 稍微寬一點
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 5, 0),
                Cursor = Cursors.Hand,
                Font = new Font(this.Font, FontStyle.Bold),
                FlatAppearance = { BorderSize = 0 },
                BackColor = isActive ? Color.CornflowerBlue : Color.LightGray,
                ForeColor = isActive ? Color.White : Color.Black
            };
        }

        private void LoadIcon()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "machine_production_history_icon.ico");
                if (File.Exists(path)) this.Icon = new Icon(path);
            }
            catch { }
        }
    }
}