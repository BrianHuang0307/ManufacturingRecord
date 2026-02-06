using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System;
using System.Windows.Forms;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;

namespace ManufacturingRecord
{
    // 這裡必須是 MachineRecordView，且要是 partial
    public partial class MachineRecordView
    {
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

        private async void _exportExcelButton_Click(object? sender, EventArgs e)
        {
            if (_excelService == null)
            {
                MessageBox.Show("ExcelService 未初始化。");
                return;
            }

            if (dgv.Rows.Count == 0 || _originalDataTable == null)
            {
                MessageBox.Show("目前畫面沒有資料可以匯出。");
                return;
            }

            string fileName = $"{_fromDateTimePicker.Text}_{_toDateTimePicker.Text}機器生產履歷.xlsx";
            string sheetName = "機器生產履歷";

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = fileName
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            _exportExcelButton.Enabled = false;
            // this.Cursor = Cursors.WaitCursor;

            try
            {
                // DataTable dt = GetDataTableFromGrid(dgv);

                DataTable dt = _originalDataTable.DefaultView.ToTable();

                await Task.Run(() =>
                {
                    // 這裡呼叫正確的方法 ExportDataToExcel
                    _excelService.ExportDataToExcel(dt, sfd.FileName, sheetName);
                });

                MessageBox.Show("機器生產履歷匯出完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("機器生產履歷匯出失敗： " + ex.Message);
            }
            finally
            {
                _exportExcelButton.Enabled = true;
                // this.Cursor = Cursors.Default;
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
        /// 將使用者輸入的字串依分隔符號切割，並組合成 (A LIKE %...% OR B LIKE %...%) 的篩選字串
        /// </summary>
        /// <param name="columnName">資料庫欄位名稱</param>
        /// <param name="inputText">使用者輸入的文字</param>
        /// <returns>RowFilter 用的篩選字串</returns>
        private string? GetMultiSearchFilter(string columnName, string inputText)
        {
            if (string.IsNullOrWhiteSpace(inputText)) return null;

            // 1. 定義多重篩選分隔分隔符號
            char[] delimiters = new char[] { ',', ' ', '、', ';' };

            // 2. 切割字串並移除空白項目
            string[] keywords = inputText.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

            if (keywords.Length == 0) return null;

            // 3. 針對每個關鍵字建立 LIKE 語句
            List<string> subFilters = new List<string>();
            foreach (string key in keywords)
            {
                // 重要：一定要處理單引號，避免 Crash
                string safeKey = key.Trim().Replace("'", "''");
                subFilters.Add($"{columnName} LIKE '%{safeKey}%'");
            }

            // 4. 用 OR 連接，並用括號包起來
            // 結果會像： (生產料件 LIKE '%A%' OR 生產料件 LIKE '%B%')
            return "(" + string.Join(" OR ", subFilters) + ")";
        }

        /// <summary>
        /// 選取資料：還原 _originalDataTable 並套用篩選 (包含機時上下限)
        /// </summary>
        private void _selectDgvButton_Click(object? sender, EventArgs e)
        {
            if (_originalDataTable == null || _originalDataTable.Rows.Count == 0)
            {
                MessageBox.Show("請先查詢資料庫取得資料。", "無資料");
                return;
            }

            // 還原資料來源
            if (dgv.DataSource != _originalDataTable)
            {
                dgv.DataSource = _originalDataTable;
            }
            ResetGridColor();

            List<string> filters = new List<string>();

            // 1. 生產料件 (支援逗號分隔)
            // 假設輸入 "A, B"，會產生 (生產料件 LIKE '%A%' OR 生產料件 LIKE '%B%')
            string? productFilter = GetMultiSearchFilter("生產料件", _productTextBox.Text);
            if (productFilter != null) filters.Add(productFilter);

            // 2. 品名簡稱
            string? abbreviationFilter = GetMultiSearchFilter("品名簡稱", _productAbbreviationTextBox.Text);
            if (abbreviationFilter != null) filters.Add(abbreviationFilter);

            // 3. 特性編碼 (支援逗號分隔)
            string? featureFilter = GetMultiSearchFilter("特性編碼", _featureTextBox.Text);
            if (featureFilter != null) filters.Add(featureFilter);

            // 4. 工序 (支援逗號分隔，需注意 Convert 語法)
            string? processFilter = GetMultiSearchFilter("Convert([工序], 'System.String')", _processTextBox.Text);
            if (processFilter != null) filters.Add(processFilter);

            // 5. 機台編號
            string? machineFilter = GetMultiSearchFilter("機台編號", _machineTextBox.Text);
            if (machineFilter != null) filters.Add(machineFilter);

            // --- [新增] 6. 投入機時(分鐘) 下限 ---
            if (decimal.TryParse(_groundTextBox.Text.Trim(), out decimal minMachineTime))
            {
                filters.Add($"[投入機時(分鐘)] >= {minMachineTime}");
            }

            // --- [新增] 7. 投入機時(分鐘) 上限 ---
            if (decimal.TryParse(_ceilTextBox.Text.Trim(), out decimal maxMachineTime))
            {
                filters.Add($"[投入機時(分鐘)] <= {maxMachineTime}");
            }

            // 組合篩選字串
            string finalFilter = filters.Count > 0 ? string.Join(" AND ", filters) : "";

            try
            {
                _originalDataTable.DefaultView.RowFilter = finalFilter;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"篩選發生錯誤: {ex.Message}");
            }

            _calculateAverageCapacityButton.Enabled = true;
            _calculateDailyCapacityButton.Enabled = true;
        }

        /// <summary>
        /// 日產能計算 (依工單分組) - 動態依據 UI 上下限篩選
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
                // 1. 取得 UI 設定的上下限 (若為空值則給予極大/極小預設值，達到不設限效果)
                decimal minLimit = 0;
                decimal maxLimit = decimal.MaxValue;

                if (!string.IsNullOrWhiteSpace(_groundTextBox.Text) && decimal.TryParse(_groundTextBox.Text.Trim(), out decimal pMin))
                    minLimit = pMin;

                if (!string.IsNullOrWhiteSpace(_ceilTextBox.Text) && decimal.TryParse(_ceilTextBox.Text.Trim(), out decimal pMax))
                    maxLimit = pMax;

                // 2. 準備資料
                DataTable dtFiltered = dtRaw.DefaultView.ToTable();

                var query = dtFiltered.AsEnumerable()
                    .Where(row => {
                        decimal val = row.IsNull("投入機時(分鐘)") ? 0 : Convert.ToDecimal(row["投入機時(分鐘)"]);
                        return val >= minLimit && val <= maxLimit;
                    })
                    // ---------------------------------------------
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

                        var machines = g.Select(r => r["機台編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedMachineIds = string.Join(",", machines);

                        decimal actualDailyCapacity = actualMachineSec > 0 ? Math.Round(86400 / actualMachineSec, 0) : 0;
                        decimal stdDailyCapacity = stdMachineSec > 0 ? Math.Round(86400 / stdMachineSec, 0) : 0;

                        return new
                        {
                            生產料件 = first["生產料件"],
                            品名簡稱 = first["品名簡稱"],
                            特性編碼 = first["特性編碼"],
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

                BindResultToGrid(query, Color.MistyRose);

                // 產生提示訊息字串
                string limitMsg = "";
                if (minLimit > 0) limitMsg += $"下限 {minLimit} ";
                if (maxLimit < decimal.MaxValue) limitMsg += $"上限 {maxLimit} ";
                if (string.IsNullOrEmpty(limitMsg)) limitMsg = "無限制";

                MessageBox.Show($"日產能計算完成！(機時條件: {limitMsg})\n共彙總 {query.Count} 筆資料。", "計算成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"計算過程中發生錯誤：\n{ex.Message}", "計算錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            _calculateDailyCapacityButton.Enabled = false;
            _calculateAverageCapacityButton.Enabled = false;
        }

        /// <summary>
        /// 平均產能計算 (依 料件+特性+工序 分組) - 動態依據 UI 上下限篩選
        /// </summary>
        private void _calculateAverageCapacityButton_Click(object? sender, EventArgs e)
        {
            DataTable? dtRaw = dgv.DataSource as DataTable;

            if (dtRaw == null || dtRaw.DefaultView.Count == 0)
            {
                MessageBox.Show("目前無資料可供計算，請先進行查詢與選取。", "無資料", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // 1. 取得 UI 設定的上下限
                decimal minLimit = 0;
                decimal maxLimit = decimal.MaxValue;

                if (!string.IsNullOrWhiteSpace(_groundTextBox.Text) && decimal.TryParse(_groundTextBox.Text.Trim(), out decimal pMin))
                    minLimit = pMin;

                if (!string.IsNullOrWhiteSpace(_ceilTextBox.Text) && decimal.TryParse(_ceilTextBox.Text.Trim(), out decimal pMax))
                    maxLimit = pMax;

                // 2. 準備資料
                DataTable dtFiltered = dtRaw.DefaultView.ToTable();

                var query = dtFiltered.AsEnumerable()
                    // --- [修改] 使用變數進行篩選 ---
                    .Where(row => {
                        decimal val = row.IsNull("投入機時(分鐘)") ? 0 : Convert.ToDecimal(row["投入機時(分鐘)"]);
                        return val >= minLimit && val <= maxLimit;
                    })
                    // -----------------------------
                    .GroupBy(row => new
                    {
                        Product = row["生產料件"].ToString(),
                        Feature = row["特性編碼"].ToString(),
                        Process = row["工序"].ToString()
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

                        var machines = g.Select(r => r["機台編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedMachineIds = string.Join(",", machines);

                        var workOrders = g.Select(r => r["工單編號"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();
                        string joinedWorkOrders = string.Join(",", workOrders);

                        decimal actualDailyCapacity = actualMachineSec > 0 ? Math.Round(86400 / actualMachineSec, 0) : 0;
                        decimal stdDailyCapacity = stdMachineSec > 0 ? Math.Round(86400 / stdMachineSec, 0) : 0;

                        return new
                        {
                            生產料件 = first["生產料件"],
                            品名簡稱 = first["品名簡稱"],
                            特性編碼 = first["特性編碼"],
                            機台編號 = joinedMachineIds,
                            工序 = first["工序"],
                            工單編號 = joinedWorkOrders,
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

                BindResultToGrid(query, Color.Honeydew);

                string limitMsg = "";
                if (minLimit > 0) limitMsg += $"下限 {minLimit} ";
                if (maxLimit < decimal.MaxValue) limitMsg += $"上限 {maxLimit} ";
                if (string.IsNullOrEmpty(limitMsg)) limitMsg = "無限制";

                MessageBox.Show($"平均產能計算完成！(機時條件: {limitMsg})\n共彙總 {query.Count} 筆資料。", "計算成功");
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
            dgv.DefaultCellStyle.BackColor = Color.Empty;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
        }
    }
}