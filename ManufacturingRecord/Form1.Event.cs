using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Windows.Forms;
using System.Data;
using System.Linq; // 必須引用 System.Linq 才能使用 GroupBy
using System.Collections.Generic;
using System.Drawing;

namespace ManufacturingRecord
{
    public partial class Form1
    {
        /* --- 開關切換 ---
        static bool FuzzySelectSwitch = true;
        */

        // --- 保存原始查詢資料 (避免計算後資料遺失) ---
        private DataTable? _originalDataTable = null;

        //  新增方法來註冊所有事件
        private void AddEventHandlers()
        {
            // 1. 查詢資料庫按鈕
            _searchDataButton.Click += _searchButton_Click;

            // 2. 匯出 Excel 按鈕
            _exportExcelButton.Click += _exportExcelButton_Click;

            // 3. 選取資料按鈕 (本機篩選)
            _selectDgvButton.Click += _selectDgvButton_Click;

            // 4. DGV 點擊事件 (填入 TextBox)
            dgv.CellClick += Dgv_CellClick;

            // 5. 模糊搜尋開關
            // _fuzzySelectCheckBox.CheckedChanged += _fuzzySelectCheckBox_CheckedChanged;

            // 5. 日產能計算按鈕
            _calculateDailyCapacityButton.Click += _calculateDailyCapacityButton_Click;

            // 6. 平均日產能計算按鈕
            _calculateAverageCapacityButton.Click += _calculateAverageCapacityButton_Click;
            // ... 其他事件
        }

        private void _searchButton_Click(object? sender, EventArgs e)
        {
            DateTime inputFromDate = _fromDateTimePicker.Value.Date;
            DateTime inputToDate = _toDateTimePicker.Value.Date;

            if (string.IsNullOrEmpty(inputFromDate.ToString()) || string.IsNullOrEmpty(inputToDate.ToString())) // || string.IsNullOrEmpty(inputProduct) || string.IsNullOrEmpty(inputFeature) || string.IsNullOrEmpty(inputProcess)
            {
                MessageBox.Show("請確保欄位輸入完整，不能有空值。", "輸入錯誤");
                return;
            }

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            var dt = _dataService.QueryMachineManufacturingResume(inputFromDate, inputToDate);

            // 保存原始資料
            _originalDataTable = dt;

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dgv.DataSource = null;
            dgv.DataSource = _originalDataTable; // 綁定原始資料
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            ResetGridColor();

            _calculateDailyCapacityButton.Enabled = true;
            _calculateAverageCapacityButton.Enabled = true;

            MessageBox.Show($"資料載入成功！共 {dt.Rows.Count} 筆記錄。", "完成");
        }

        private void _exportExcelButton_Click(object? sender, EventArgs e)
        {
            // 呼叫 Excel 匯出 Service，傳入 dgv.DataSource (DataTable)
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("目前畫面沒有資料可以匯出。");
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = _fromDateTimePicker.Text + "_" + _toDateTimePicker.Text + "機器生產履歷.xlsx"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                //  直接匯出當前的 DataGridView 狀態
                _excelService.ExportGridToExcel(dgv, sfd.FileName);
                MessageBox.Show("匯出完成！（以目前 Grid 顯示為準）");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
        }

        // 點擊 DataGridView 儲存格，將內容帶入對應的 TextBox
        private void Dgv_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            // 排除點擊標題列或無效區域
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // 取得目前點擊的欄位標題 (HeaderText)
            string headerText = dgv.Columns[e.ColumnIndex].HeaderText;

            // 取得儲存格的值 (處理 null)
            string cellValue = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";

            // 根據標題填入對應 TextBox
            switch (headerText)
            {
                case "生產料件":
                    _productTextBox.Text = cellValue;
                    break;
                case "品名簡稱":
                    _productAbbreviationTextBox.Text = cellValue;
                    break;
                case "特性編碼":
                    _featureTextBox.Text = cellValue;
                    break;
                case "工序":
                    _processTextBox.Text = cellValue;
                    break;
                case "機台編號":
                    _machineTextBox.Text = cellValue;
                    break;
                    // 其他欄位點擊不做反應
            }
        }

        /// <summary>
        /// 選取資料：還原 _originalDataTable 並套用篩選
        /// </summary>
        private void _selectDgvButton_Click(object? sender, EventArgs e)
        {
            // 檢查是否有原始資料可供篩選
            if (_originalDataTable == null || _originalDataTable.Rows.Count == 0)
            {
                MessageBox.Show("請先查詢資料庫取得資料。", "無資料");
                return;
            }

            // --- 關鍵修改：還原資料來源 ---
            // 無論目前 DGV 顯示的是什麼 (可能是計算結果)，都先切換回原始資料
            if (dgv.DataSource != _originalDataTable)
            {
                dgv.DataSource = _originalDataTable;
            }

            // --- 重置顏色 (因為切回了原始資料) ---
            ResetGridColor();

            // 收集篩選條件
            List<string> filters = new List<string>();

            // 1. 生產料件 (使用 LIKE 做模糊搜尋，若要精確搜尋請改用 = '{value}')
            // Replace("'", "''") 是為了防止篩選字串中有單引號導致語法錯誤
            if (!string.IsNullOrWhiteSpace(_productTextBox.Text))
            {
                // if (FuzzySelectSwitch)
                filters.Add($"生產料件 LIKE '%{_productTextBox.Text.Trim().Replace("'", "''")}%'");
                /*
                else
                    // 精準搜尋：補上單引號
                    filters.Add($"[生產料件] = '{_productTextBox.Text.Trim().Replace("'", "''")}'");
                */
            }

            // 2. 品名簡稱
            if (!string.IsNullOrWhiteSpace(_productAbbreviationTextBox.Text))
            {
                // if (FuzzySelectSwitch)
                filters.Add($"品名簡稱 LIKE '%{_productAbbreviationTextBox.Text.Trim().Replace("'", "''")}%'");
                /*
                else
                    filters.Add($"[品名簡稱] = '{_productAbbreviationTextBox.Text.Trim().Replace("'", "''")}'");
                */
            }

            // 3. 特性編碼
            if (!string.IsNullOrWhiteSpace(_featureTextBox.Text))
            {
                // if (FuzzySelectSwitch)
                    filters.Add($"特性編碼 LIKE '%{_featureTextBox.Text.Trim().Replace("'", "''")}%'");
                /*
                else
                    filters.Add($"[特性編碼] = '{_featureTextBox.Text.Trim().Replace("'", "''")}'");
                */
            }

            // 4. 工序 (因為是 int，需先轉成 String 才能用 LIKE)
            if (!string.IsNullOrWhiteSpace(_processTextBox.Text))
            {
                // if (FuzzySelectSwitch)
                filters.Add($"Convert([工序], 'System.String') LIKE '%{_processTextBox.Text.Trim().Replace("'", "''")}%'");
                /*
                else
                {
                    // 精準搜尋：如果是數字則不加引號，但要確保轉型成功
                    if (int.TryParse(_processTextBox.Text.Trim(), out int processValue))
                    {
                        filters.Add($"[工序] = {processValue}");
                    }
                    else
                    {
                        // 若輸入非數字但資料庫為數字欄位，這裡可選擇不加入條件或加入一個必失敗條件
                        // 為避免報錯，此處不做處理，視為無法匹配該條件
                    }
                    
                }
                */
            }

            // 5. 機台編號
            if (!string.IsNullOrWhiteSpace(_machineTextBox.Text))
            {
                filters.Add($"機台編號 LIKE '%{_machineTextBox.Text.Trim().Replace("'", "''")}%'");
            }

            // 組合篩選字串 (用 AND 連接)
            string finalFilter = filters.Count > 0 ? string.Join(" AND ", filters) : "";

            try
            {
                // 設定 RowFilter
                // 若 finalFilter 為空字串，RowFilter 會自動顯示所有資料
                _originalDataTable.DefaultView.RowFilter = finalFilter;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"篩選發生錯誤: {ex.Message}");
            }

            _calculateAverageCapacityButton.Enabled = true;
            _calculateDailyCapacityButton.Enabled = true;
        }

        /*
        private void _fuzzySelectCheckBox_CheckedChanged(object? sender, EventArgs e)
        {
            if (_fuzzySelectCheckBox.Checked == true)
            {
                _calculateDailyCapacityButton.Enabled = false;
                FuzzySelectSwitch = true;
            }
            else
            {
                _calculateDailyCapacityButton.Enabled = true;
                FuzzySelectSwitch = false;
            }
        }
        */

        
        /// <summary>
        /// 日產能計算 (依工單分組)
        /// </summary>
        private void _calculateDailyCapacityButton_Click(object? sender, EventArgs e)
        {
            DataTable? dtRaw = dgv.DataSource as DataTable;

            if (dtRaw == null || dtRaw.DefaultView.Count == 0)
            {
                MessageBox.Show("目前無資料可供計算，請先進行查詢與選取。", "無資料", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                DataTable dtFiltered = dtRaw.DefaultView.ToTable();

                var query = dtFiltered.AsEnumerable()
                    .GroupBy(row => new
                    {
                        Product = row["生產料件"].ToString(),
                        Feature = row["特性編碼"].ToString(),
                        Process = row["工序"].ToString(),
                        WorkOrder = row["工單編號"].ToString()
                    })
                    .Select(g =>
                    {
                        decimal goodQty = g.Sum(r => r.IsNull("良品轉出數量") ? 0 : Convert.ToDecimal(r["良品轉出數量"]));
                        decimal scrapQty = g.Sum(r => r.IsNull("當站報廢數量") ? 0 : Convert.ToDecimal(r["當站報廢數量"]));
                        decimal inputLaborMin = g.Sum(r => r.IsNull("投入工時(分鐘)") ? 0 : Convert.ToDecimal(r["投入工時(分鐘)"]));
                        decimal inputMachineMin = g.Sum(r => r.IsNull("投入機時(分鐘)") ? 0 : Convert.ToDecimal(r["投入機時(分鐘)"]));

                        decimal totalQty = goodQty + scrapQty;
                        decimal actualLaborSec = totalQty == 0 ? 0 : Math.Round((inputLaborMin * 60) / totalQty, 2);
                        decimal actualMachineSec = totalQty == 0 ? 0 : Math.Round((inputMachineMin * 60) / totalQty, 2);

                        DataRow first = g.First();
                        decimal stdLaborSec = (first["標準人工工時(秒/pcs)"] == DBNull.Value) ? 0 : Convert.ToDecimal(first["標準人工工時(秒/pcs)"]);
                        decimal stdMachineSec = (first["標準機器工時(秒/pcs)"] == DBNull.Value) ? 0 : Convert.ToDecimal(first["標準機器工時(秒/pcs)"]);

                        // 合併機台
                        var machines = g.Select(r => r["機台編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedMachineIds = string.Join(",", machines);

                        // 計算日產能
                        decimal actualDailyCapacity = actualMachineSec > 0 ? Math.Round(86400 / actualMachineSec, 0) : 0;
                        decimal stdDailyCapacity = stdMachineSec > 0 ? Math.Round(86400 / stdMachineSec, 0) : 0;

                        return new
                        {
                            生產料件 = first["生產料件"],
                            品名簡稱 = first["品名簡稱"],
                            特性編碼 = first["特性編碼"],
                            // 工單編號 = first["工單編號"],
                            機台編號 = joinedMachineIds,
                            工序 = first["工序"],
                            工單編號 = first["工單編號"],
                            總數量 = totalQty,
                            實際人工工時 = actualLaborSec,
                            標準人工工時 = stdLaborSec,
                            實際機器工時 = actualMachineSec,
                            標準機器工時 = stdMachineSec,
                            實際日產能 = actualDailyCapacity,
                            標準日產能 = stdDailyCapacity,
                            良品總數 = goodQty,
                            報廢總數 = scrapQty,
                            投入工時分 = inputLaborMin,
                            投入機時分 = inputMachineMin
                        };
                    }).ToList();

                // 呼叫綁定方法，並傳入對應的顏色 (MistyRose)
                BindResultToGrid(query, Color.MistyRose);
                MessageBox.Show($"日產能計算完成！已將 {dtFiltered.Rows.Count} 筆原始資料合併為 {query.Count} 筆彙總資料。" +
                    $"\n日產能以生產料件+特性編碼+工序+工單編號作為最小單位進行加總", "計算成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"計算過程中發生錯誤：\n{ex.Message}", "計算錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            _calculateDailyCapacityButton.Enabled = false;
            _calculateAverageCapacityButton.Enabled = false;
        }

        /// <summary>
        /// 平均產能計算 (依 料件+特性+工序 分組，合併工單)
        /// </summary>
        private void _calculateAverageCapacityButton_Click(object? sender, EventArgs e)
        {
            // 1. 檢查目前 Grid 是否有資料
            DataTable? dtRaw = dgv.DataSource as DataTable;

            if (dtRaw == null || dtRaw.DefaultView.Count == 0)
            {
                MessageBox.Show("目前無資料可供計算，請先進行查詢與選取。", "無資料", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // 取得目前篩選後的資料表
                DataTable dtFiltered = dtRaw.DefaultView.ToTable();

                // 2. 分組：只使用 生產料件 + 特性編碼 + 工序 (忽略工單編號)
                var query = dtFiltered.AsEnumerable()
                    .GroupBy(row => new
                    {
                        Product = row["生產料件"].ToString(),
                        Feature = row["特性編碼"].ToString(),
                        Process = row["工序"].ToString()
                    })
                    .Select(g =>
                    {
                        // 彙總所有工單的數據
                        decimal goodQty = g.Sum(r => r.IsNull("良品轉出數量") ? 0 : Convert.ToDecimal(r["良品轉出數量"]));
                        decimal scrapQty = g.Sum(r => r.IsNull("當站報廢數量") ? 0 : Convert.ToDecimal(r["當站報廢數量"]));
                        decimal inputLaborMin = g.Sum(r => r.IsNull("投入工時(分鐘)") ? 0 : Convert.ToDecimal(r["投入工時(分鐘)"]));
                        decimal inputMachineMin = g.Sum(r => r.IsNull("投入機時(分鐘)") ? 0 : Convert.ToDecimal(r["投入機時(分鐘)"]));

                        decimal totalQty = goodQty + scrapQty;

                        // 計算平均秒/pcs (總投入時間 / 總數量)
                        decimal actualLaborSec = totalQty == 0 ? 0 : Math.Round((inputLaborMin * 60) / totalQty, 2);
                        decimal actualMachineSec = totalQty == 0 ? 0 : Math.Round((inputMachineMin * 60) / totalQty, 2);

                        DataRow first = g.First();
                        decimal stdLaborSec = (first["標準人工工時(秒/pcs)"] == DBNull.Value) ? 0 : Convert.ToDecimal(first["標準人工工時(秒/pcs)"]);
                        decimal stdMachineSec = (first["標準機器工時(秒/pcs)"] == DBNull.Value) ? 0 : Convert.ToDecimal(first["標準機器工時(秒/pcs)"]);

                        // --- 合併字串 ---
                        // 1. 機台編號合併
                        var machines = g.Select(r => r["機台編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedMachineIds = string.Join(",", machines);

                        // 2. 工單編號合併 (新規則)
                        var workOrders = g.Select(r => r["工單編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedWorkOrders = string.Join(",", workOrders);

                        // 計算日產能 (依據平均速率)
                        decimal actualDailyCapacity = actualMachineSec > 0 ? Math.Round(86400 / actualMachineSec, 0) : 0;
                        decimal stdDailyCapacity = stdMachineSec > 0 ? Math.Round(86400 / stdMachineSec, 0) : 0;

                        return new
                        {
                            生產料件 = first["生產料件"],
                            品名簡稱 = first["品名簡稱"],
                            特性編碼 = first["特性編碼"],
                            機台編號 = joinedMachineIds, // 顯示合併後的機台
                            工序 = first["工序"],
                            工單編號 = joinedWorkOrders, // 顯示合併後的工單

                            總數量 = totalQty,

                            實際人工工時 = actualLaborSec,
                            標準人工工時 = stdLaborSec,
                            實際機器工時 = actualMachineSec,
                            標準機器工時 = stdMachineSec,

                            實際日產能 = actualDailyCapacity,
                            標準日產能 = stdDailyCapacity,

                            良品總數 = goodQty,
                            報廢總數 = scrapQty,
                            投入工時分 = inputLaborMin,
                            投入機時分 = inputMachineMin
                        };
                    })
                    .ToList();

                // 呼叫綁定方法，並傳入對應的顏色 (Honeydew)
                BindResultToGrid(query, Color.Honeydew);
                MessageBox.Show($"平均產能計算完成！已將 {dtFiltered.Rows.Count} 筆原始資料合併為 {query.Count} 筆彙總資料。" +
                    $"\n平均產能以生產料件+特性編碼+工序作為最小單位進行加總", "計算成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"計算過程中發生錯誤：\n{ex.Message}", "計算錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            _calculateDailyCapacityButton.Enabled = false;
            _calculateAverageCapacityButton.Enabled = false;
        }
        
        // 共用的綁定 Grid 方法，避免重複代碼
        // 修改後的綁定方法，增加 bgColor 參數
        private void BindResultToGrid(dynamic queryList, Color bgColor)
        {
            DataTable resultTable = new DataTable();
            resultTable.Columns.Add("生產料件");
            resultTable.Columns.Add("品名簡稱");
            resultTable.Columns.Add("特性編碼");
            resultTable.Columns.Add("機台編號");
            resultTable.Columns.Add("工序");
            resultTable.Columns.Add("工單編號");
            resultTable.Columns.Add("總數量", typeof(decimal));

            resultTable.Columns.Add("實際人工工時(秒/pcs)", typeof(decimal));
            resultTable.Columns.Add("標準人工工時(秒/pcs)", typeof(decimal));
            resultTable.Columns.Add("實際機器工時(秒/pcs)", typeof(decimal));
            resultTable.Columns.Add("標準機器工時(秒/pcs)", typeof(decimal));

            resultTable.Columns.Add("實際日產能(pcs/天)", typeof(decimal));
            resultTable.Columns.Add("標準日產能(pcs/天)", typeof(decimal));

            resultTable.Columns.Add("良品總數", typeof(decimal));
            resultTable.Columns.Add("報廢總數", typeof(decimal));
            resultTable.Columns.Add("投入工時(分)", typeof(decimal));
            resultTable.Columns.Add("投入機時(分)", typeof(decimal));

            foreach (var item in queryList)
            {
                resultTable.Rows.Add(
                    item.生產料件,
                    item.品名簡稱,
                    item.特性編碼,
                    item.機台編號,
                    item.工序,
                    item.工單編號,
                    item.總數量,
                    item.實際人工工時,
                    item.標準人工工時,
                    item.實際機器工時,
                    item.標準機器工時,
                    item.實際日產能,
                    item.標準日產能,
                    item.良品總數,
                    item.報廢總數,
                    item.投入工時分,
                    item.投入機時分
                );
            }

            dgv.DataSource = resultTable;

            // --- 設定背景顏色 ---
            dgv.DefaultCellStyle.BackColor = bgColor;
            // 讓交替行也保持相同色調 (或者你可以刪除這行讓它變回預設)
            dgv.AlternatingRowsDefaultCellStyle.BackColor = bgColor;
        }

        // 輔助方法：重置 Grid 顏色回預設 (AliceBlue/White)
        private void ResetGridColor()
        {
            // 這裡設為 Empty 或 White 或 AliceBlue 皆可
            dgv.DefaultCellStyle.BackColor = Color.Empty;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
        }
    }
}