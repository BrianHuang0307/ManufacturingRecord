using DocumentFormat.OpenXml.Spreadsheet;
using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace ManufacturingRecord
{
    public partial class ProductErrorCodeView
    {
        // --- 保存原始查詢資料 ---
        private DataTable? _originalDataTable = null;

        private class AnalysisResult
        {
            public string 生產料件 { get; set; } = "";
            public string 特性編碼 { get; set; } = "";
            public string 工序 { get; set; } = "";
            public string 異常項目 { get; set; } = ""; // 顯示【原因】或 └ [備註]
            public decimal 報廢數量 { get; set; }
            public decimal 完工入庫數量 { get; set; }
            public decimal 異常比例 { get; set; }
            public string RowType { get; set; } = ""; // 用來分辨 Parent 或 Child
        }

        private void AddEventHandlers()
        {
            // 從資料庫獲取資料
            _searchDataButton.Click += _searchButton_Click;

            // 將當前datagridview的資料會出成Excel檔案
            _exportExcelButton.Click += _exportExcelButton_Click;

            // 根據使用者TextBox欄位的值進行模糊搜尋
            _selectDgvButton.Click += _selectDgvButton_Click;

            // 點擊資料欄位，將資料映射到使用者TextBox輸入欄位
            dgv.CellClick += Dgv_CellClick;

            // 點擊計算按鈕，進行分組彙總運算
            _calculateButton.Click += _calculateButton_Click;
            dgv.CellClick += Dgv_CellClick_CollapseExpand; // 註冊摺疊點擊事件
        }

        private void _searchButton_Click(object? sender, EventArgs e)
        {
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            try
            {
                // [修正] 原本的程式碼沒有呼叫 _dataService，只是把 null 指派回去
                // 這裡呼叫 DataService 取得資料
                var dt = _dataService.QueryProductErrorCode();

                _originalDataTable = dt;

                dgv.DataSource = null;
                dgv.DataSource = _originalDataTable;

                MessageBox.Show($"資料載入成功！共 {dt.Rows.Count} 筆記錄。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查詢失敗: {ex.Message}", "錯誤");
            }
            finally
            {
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            }
        }

        private async void _exportExcelButton_Click(object? sender, EventArgs e)
        {
            if (_excelService == null || dgv.Rows.Count == 0) return;

            string fileName = $"產品錯誤代碼統計_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            using var sfd = new SaveFileDialog { Filter = "Excel Files|*.xlsx", FileName = fileName };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            _exportExcelButton.Enabled = false;
            try
            {
                DataTable dtExport;

                // [判斷] 如果有 "RowType" 欄位，代表目前是「計算統計模式」，需要進行扁平化處理以便 Excel 分析
                if (dgv.Columns.Contains("RowType"))
                {
                    dtExport = GetFlatTableForAnalysis(dgv);
                }
                else
                {
                    // 否則就是一般的「原始資料模式」，照舊匯出
                    dtExport = GetDataTableFromGrid(dgv);
                }

                await Task.Run(() => _excelService.ExportDataToExcel(dtExport, sfd.FileName, "產品錯誤代碼"));
                MessageBox.Show("匯出完成！\n已自動轉換為適合樞紐分析的格式（填補空白欄位並拆分原因備註）。");
            }
            catch (Exception ex) { MessageBox.Show("匯出失敗： " + ex.Message); }
            finally { _exportExcelButton.Enabled = true; }
        }

        /*
        private async void _exportExcelButton_Click(object? sender, EventArgs e)
        {
            if (_excelService == null)
            {
                MessageBox.Show("ExcelService 未初始化。");
                return;
            }

            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("目前畫面沒有資料可以匯出。");
                return;
            }

            // [修正] 定義檔名與 Sheet 名稱
            string sheetName = "產品錯誤代碼";
            string fileName = $"產品錯誤代碼統計_{DateTime.Now:yyyyMMdd}.xlsx";

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
                // 1. 在 UI 執行緒先執行 GetDataTableFromGrid
                //    這樣才能抓到「計算後」的階層資料，以及「當下」的排序與篩選
                DataTable dtExport = GetDataTableFromGrid(dgv);

                // 2. 將抓好的資料丟給 Service 進行匯出
                await Task.Run(() =>
                {
                    _excelService.ExportDataToExcel(dtExport, sfd.FileName, sheetName);
                });

                MessageBox.Show("匯出完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
            finally
            {
                _exportExcelButton.Enabled = true;
            }
        }
        */

        private void Dgv_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            string headerText = dgv.Columns[e.ColumnIndex].HeaderText;
            string cellValue = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";

            // [注意] 這裡的 Case 名稱對應 SQL 裡的 SELECT AS "別名"
            switch (headerText)
            {
                case "生產料件":
                    // if (_productTextBox != null) 
                    _productTextBox.Text = cellValue;
                    break;
                /*
                case "品名簡稱":
                    // if (_workOrderTextBox != null) 
                    _productAbbreviationTextBox.Text = cellValue;
                    break;
                */
                case "特性編碼":
                    // if (_machineTextBox != null) 
                    _featureTextBox.Text = cellValue;
                    break;
                case "工序":
                    // if (_errorCodeTextBox != null)
                    _processTextBox.Text = cellValue;
                    break;
                    // 若有其他欄位需要對應，加在這裡
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

        private void _selectDgvButton_Click(object? sender, EventArgs e)
        {
            if (_originalDataTable == null || _originalDataTable.Rows.Count == 0)
            {
                MessageBox.Show("請先查詢資料庫取得資料。", "無資料");
                return;
            }

            // 重置 DataSource 以確保篩選是基於原始資料
            if (dgv.DataSource != _originalDataTable) dgv.DataSource = _originalDataTable;

            List<string> filters = new List<string>();

            // 1. 生產料件 (支援逗號分隔)
            // 假設輸入 "A, B"，會產生 (生產料件 LIKE '%A%' OR 生產料件 LIKE '%B%')
            string? productFilter = GetMultiSearchFilter("生產料件", _productTextBox.Text);
            if (productFilter != null) filters.Add(productFilter);

            /*
            if (!string.IsNullOrWhiteSpace(_productAbbreviationTextBox.Text))
                filters.Add($"品名簡稱 LIKE '%{_productAbbreviationTextBox.Text.Trim().Replace("'", "''")}%'");
            */

            // 2. 特性編碼 (支援逗號分隔)
            string? featureFilter = GetMultiSearchFilter("特性編碼", _featureTextBox.Text);
            if (featureFilter != null) filters.Add(featureFilter);

            // 3. 工序 (支援逗號分隔，需注意 Convert 語法)
            string? processFilter = GetMultiSearchFilter("Convert([工序], 'System.String')", _processTextBox.Text);
            if (processFilter != null) filters.Add(processFilter);

            try
            {
                string finalFilter = filters.Count > 0 ? string.Join(" AND ", filters) : "";
                _originalDataTable.DefaultView.RowFilter = finalFilter;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"篩選發生錯誤: {ex.Message}\n請檢查輸入字元是否合法。");
            }
        }

        private DataTable GetDataTableFromGrid(DataGridView grid)
        {
            var dt = new DataTable();

            var sortedColumns = grid.Columns.Cast<DataGridViewColumn>()
                                            .OrderBy(c => c.DisplayIndex)
                                            .Where(c => c.Visible)
                                            .ToList();

            foreach (var col in sortedColumns)
            {
                dt.Columns.Add(col.HeaderText);
            }

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;

                DataRow dr = dt.NewRow();
                for (int i = 0; i < sortedColumns.Count; i++)
                {
                    var value = row.Cells[sortedColumns[i].Name].Value;
                    dr[i] = value == null ? "" : value.ToString();
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        // ---------------------------------------------------------
        //  核心計算邏輯 (含 Child 合併與摺疊初始化)
        // ---------------------------------------------------------
        private void _calculateButton_Click(object? sender, EventArgs e)
        {
            if (dgv.DataSource == null)
            {
                MessageBox.Show("目前無資料可計算，請先載入並選取資料。", "提示");
                return;
            }

            try
            {
                // 1. 取得資料來源 (優先使用原始資料的 Filter 結果)
                DataTable dtSource = (_originalDataTable != null)
                    ? _originalDataTable.DefaultView.ToTable()
                    : (DataTable)dgv.DataSource;

                if (dtSource.Rows.Count == 0) return;

                // 2. 預計算「總分母」：每個 產品/特性/工序 的總完工數量
                // (邏輯：同工序內的工單完工數只加總一次，避免重複計算)
                var productionTotals = dtSource.AsEnumerable()
                    .GroupBy(row => new
                    {
                        Product = row.Field<string>("生產料件") ?? "",
                        Feature = row.Field<string>("特性編碼") ?? "",
                        Process = row["工序"]?.ToString() ?? ""
                    })
                    .ToDictionary(
                        g => g.Key,
                        g => g.GroupBy(r => r.Field<string>("工單編號"))
                              .Sum(woGroup => woGroup.First().IsNull("完工入庫數量")
                                  ? 0 : Convert.ToDecimal(woGroup.First()["完工入庫數量"]))
                    );

                // 3. 建立結果列表
                var finalResults = new List<AnalysisResult>();

                // 第一層：依 異常原因 分組 (Parent)
                var parentGroups = dtSource.AsEnumerable()
                    .GroupBy(row => new
                    {
                        Product = row.Field<string>("生產料件") ?? "",
                        Feature = row.Field<string>("特性編碼") ?? "",
                        Process = row["工序"]?.ToString() ?? "",
                        Reason = row.Field<string>("異常原因")
                    })
                    // 排序：產品 -> 工序 -> 異常數量多到少
                    .OrderBy(x => x.Key.Product)
                    .ThenBy(x => x.Key.Feature)
                    .ThenBy(x => x.Key.Process)
                    .ThenByDescending(x => x.Sum(r => r.IsNull("數量") ? 0 : Convert.ToDecimal(r["數量"])));

                foreach (var pGroup in parentGroups)
                {
                    var pKey = new 
                    { 
                        Product = pGroup.Key.Product, 
                        Feature = pGroup.Key.Feature, 
                        Process = pGroup.Key.Process ?? ""};
                    decimal totalFinished = productionTotals.ContainsKey(pKey) ? productionTotals[pKey] : 0;

                    // 計算 Parent 總報廢數
                    decimal pScrapQty = pGroup.Sum(r => r.IsNull("數量") ? 0 : Convert.ToDecimal(r["數量"]));

                    // --- 加入 Parent 列 ---
                    finalResults.Add(new AnalysisResult
                    {
                        生產料件 = pGroup.Key.Product,
                        特性編碼 = pGroup.Key.Feature,
                        工序 = pGroup.Key.Process ?? "",
                        // 預設加上 [-] 表示展開狀態
                        異常項目 = "[-] 【原因】" + (string.IsNullOrWhiteSpace(pGroup.Key.Reason) ? "未註記" : pGroup.Key.Reason),
                        報廢數量 = pScrapQty,
                        完工入庫數量 = totalFinished,
                        異常比例 = totalFinished == 0 ? 0 : Math.Round(pScrapQty / totalFinished, 8),
                        RowType = "Parent"
                    });

                    // --- 處理 Child (異常備註) ---
                    // 1. 篩選有備註的資料
                    // 2. GroupBy 備註名稱 (將相同的 L50.32NG 合併)
                    // 3. Sum 數量
                    var childList = pGroup
                        .Where(r => !r.IsNull("異常備註") && !string.IsNullOrWhiteSpace(r.Field<string>("異常備註")))
                        .GroupBy(r => r.Field<string>("異常備註")?.Trim()) // Trim 避免空白差異
                        .Select(g => new
                        {
                            RemarkText = g.Key,
                            TotalScrapQty = g.Sum(r => r.IsNull("數量") ? 0 : Convert.ToDecimal(r["數量"]))
                        })
                        .OrderByDescending(x => x.TotalScrapQty) // 數量多的備註排上面
                        .ToList();

                    foreach (var childItem in childList)
                    {
                        // 加入 Child 列
                        finalResults.Add(new AnalysisResult
                        {
                            生產料件 = "", // 留空製造縮排視覺感
                            特性編碼 = "",
                            工序 = "",
                            異常項目 = "      └ " + childItem.RemarkText,
                            報廢數量 = childItem.TotalScrapQty,
                            完工入庫數量 = totalFinished, // 分母維持一樣，方便看佔比
                            異常比例 = totalFinished == 0 ? 0 : Math.Round(childItem.TotalScrapQty / totalFinished, 8),
                            RowType = "Child"
                        });
                    }
                }

                // 4. 綁定資料
                dgv.DataSource = null;
                dgv.DataSource = ToDataTable(finalResults);

                // 5. 設定樣式 (粗體、顏色)
                FormatGridStyle();

                // 6. [重要] 停用排序功能，防止樹狀結構被破壞
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                MessageBox.Show($"計算完成！\n已合併重複備註。\n點擊「異常項目」前的符號可摺疊/展開。", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"計算發生錯誤: {ex.Message}", "錯誤");
            }
        }

        // ---------------------------------------------------------
        //  UI 互動：摺疊與展開
        // ---------------------------------------------------------
        private void Dgv_CellClick_CollapseExpand(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // 確保有 RowType 欄位 (計算模式下才有)
            if (!dgv.Columns.Contains("RowType")) return;

            var currentRow = dgv.Rows[e.RowIndex];
            var typeValue = currentRow.Cells["RowType"].Value?.ToString();

            // 只處理 Parent 的點擊
            if (typeValue != "Parent") return;

            // 判斷點擊的是不是「異常項目」欄位
            if (dgv.Columns[e.ColumnIndex].HeaderText != "異常項目") return;

            var cell = currentRow.Cells["異常項目"];
            string currentText = cell.Value?.ToString() ?? "";
            bool isExpanded = currentText.Contains("[-]");

            if (isExpanded)
            {
                // 執行摺疊
                cell.Value = currentText.Replace("[-]", "[+]");
                ToggleChildrenVisibility(e.RowIndex, false);
            }
            else if (currentText.Contains("[+]"))
            {
                // 執行展開
                cell.Value = currentText.Replace("[+]", "[-]");
                ToggleChildrenVisibility(e.RowIndex, true);
            }
        }

        private void ToggleChildrenVisibility(int parentIndex, bool visible)
        {
            // 從 Parent 的下一行開始往下找
            for (int i = parentIndex + 1; i < dgv.Rows.Count; i++)
            {
                var row = dgv.Rows[i];
                var type = row.Cells["RowType"].Value?.ToString();

                // 遇到下一個 Parent 就停止
                if (type == "Parent") break;

                // 設定 Child 的可見度
                if (type == "Child")
                {
                    // 若使用 CurrencyManager 管理 DataBinding，直接設 Visible 即可生效
                    row.Visible = visible;
                }
            }
        }

        private void FormatGridStyle()
        {
            // 隱藏輔助欄位
            if (dgv.Columns.Contains("RowType")) dgv.Columns["RowType"].Visible = false;

            // 設定百分比格式
            if (dgv.Columns.Contains("異常比例"))
            {
                dgv.Columns["異常比例"].DefaultCellStyle.Format = "P4"; // P2 = 12.3456%
            }

            // 遍歷設定顏色
            foreach (DataGridViewRow row in dgv.Rows)
            {
                var typeCell = row.Cells["RowType"].Value?.ToString();

                if (typeCell == "Parent")
                {
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240); // 淺灰
                    row.DefaultCellStyle.Font = new System.Drawing.Font(dgv.Font, FontStyle.Bold);
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
                }
                else if (typeCell == "Child")
                {
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.White;
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.DimGray;
                }
            }
        }

        // ---------------------------------------------------------
        //  工具方法：List 轉 DataTable
        // ---------------------------------------------------------
        private DataTable ToDataTable<T>(IEnumerable<T> items)
        {
            DataTable dt = new DataTable();
            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (var item in items)
            {
                var values = new object[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null) ?? DBNull.Value;
                }
                dt.Rows.Add(values);
            }
            return dt;
        }

        // 為了 Excel 分析設計的扁平化匯出方法
        private DataTable GetFlatTableForAnalysis(DataGridView grid)
        {
            var dt = new DataTable();

            // 定義 Excel 欄位 (適合樞紐分析的格式)
            dt.Columns.Add("生產料件");
            dt.Columns.Add("特性編碼");
            dt.Columns.Add("工序");
            dt.Columns.Add("異常原因");  // 拆分出來
            dt.Columns.Add("異常備註");  // 拆分出來
            dt.Columns.Add("報廢數量", typeof(decimal));
            dt.Columns.Add("完工入庫數量", typeof(decimal));
            dt.Columns.Add("異常比例", typeof(string)); // 匯出成字串保留百分比格式，或用 decimal 也可以
            dt.Columns.Add("資料層級"); // 標記這是 "總計" 還是 "細項"，避免 User 樞紐時重複加總

            // 用來暫存 Parent 的資訊 (向下填補用)
            string lastProduct = "";
            string lastFeature = "";
            string lastProcess = "";
            string lastReason = ""; // 暫存異常原因

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;

                // 取得隱藏的 RowType
                string rowType = row.Cells["RowType"].Value?.ToString() ?? "";
                string itemText = row.Cells["異常項目"].Value?.ToString() ?? "";

                // 數值處理
                decimal scrapQty = Convert.ToDecimal(row.Cells["報廢數量"].Value ?? 0);
                decimal finishQty = Convert.ToDecimal(row.Cells["完工入庫數量"].Value ?? 0);

                // 取得異常比例 (保留 Grid 上的百分比文字，或者重算)
                string rateText = row.Cells["異常比例"].FormattedValue?.ToString() ?? "";

                DataRow dr = dt.NewRow();

                if (rowType == "Parent")
                {
                    // --- 處理 Parent (分類總計) ---

                    // 1. 更新暫存變數 (供 Child 使用)
                    lastProduct = row.Cells["生產料件"].Value?.ToString() ?? "";
                    lastFeature = row.Cells["特性編碼"].Value?.ToString() ?? "";
                    lastProcess = row.Cells["工序"].Value?.ToString() ?? "";

                    // 2. 清洗異常原因文字 (移除 "[-] 【原因】" 等符號)
                    // 假設格式固定為 "[-] 【原因】 尺寸異常" 或 "[+] 【原因】 尺寸異常"
                    // 抓 "【原因】" 後面的字
                    int splitIndex = itemText.IndexOf("【原因】");
                    lastReason = (splitIndex >= 0) ? itemText.Substring(splitIndex + 4).Trim() : itemText;

                    // 3. 填入 DataRow
                    dr["生產料件"] = lastProduct;
                    dr["特性編碼"] = lastFeature;
                    dr["工序"] = lastProcess;
                    dr["異常原因"] = lastReason;
                    dr["異常備註"] = ""; // Parent 行不含備註
                    dr["報廢數量"] = scrapQty;
                    dr["完工入庫數量"] = finishQty;
                    dr["異常比例"] = rateText;
                    dr["資料層級"] = "異常原因";
                }
                else if (rowType == "Child")
                {
                    // --- 處理 Child (詳細備註) ---

                    // 自動填補 (Fill Down) - 使用 Parent 的資料
                    dr["生產料件"] = lastProduct;
                    dr["特性編碼"] = lastFeature;
                    dr["工序"] = lastProcess;
                    dr["異常原因"] = lastReason; // 繼承 Parent 的原因

                    // int splitIndex = itemText.IndexOf("└");
                    // string remarkClean = (splitIndex >= 0) ? itemText.Substring(splitIndex + 1).Trim() : itemText;
                    string remarkClean = itemText;
                    dr["異常備註"] = remarkClean;

                    dr["報廢數量"] = scrapQty;
                    dr["完工入庫數量"] = finishQty; // 分母相同
                    dr["異常比例"] = rateText;
                    dr["資料層級"] = "異常備註";
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}
