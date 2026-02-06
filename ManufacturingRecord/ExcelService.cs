using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ManufacturingRecord.Service
{
    public class ExcelService : IExcelService
    {
        public void ExportDataToExcel(DataTable dataTable, string path, string sheetName)
        {
            using var wb = new XLWorkbook();

            // Excel 單一分頁最大列數限制 (保留一點緩衝給標題列)
            const int MaxRowsPerSheet = 1000000;

            // 如果資料量小於限制，直接寫入 (原本的邏輯)
            if (dataTable.Rows.Count <= MaxRowsPerSheet)
            {
                var ws = wb.AddWorksheet(sheetName);
                ws.Cell(1, 1).InsertTable(dataTable);
                // 資料量少才自動調整欄寬，不然會很慢
                if (dataTable.Rows.Count < 50000) ws.Columns().AdjustToContents();
            }
            else
            {
                // --- 資料量過大，進行自動分頁 ---
                int totalRows = dataTable.Rows.Count;
                int pageCount = (int)Math.Ceiling((double)totalRows / MaxRowsPerSheet);

                for (int i = 0; i < pageCount; i++)
                {
                    // 建立分頁名稱：機器生產履歷_1, 機器生產履歷_2...
                    string currentSheetName = $"{sheetName}_{i + 1}";
                    var ws = wb.AddWorksheet(currentSheetName);

                    // 取得目前分頁的資料片段
                    var rows = dataTable.AsEnumerable()
                        .Skip(i * MaxRowsPerSheet)
                        .Take(MaxRowsPerSheet);

                    // 建立臨時 DataTable 寫入
                    if (rows.Any())
                    {
                        DataTable chunkTable = rows.CopyToDataTable();
                        ws.Cell(1, 1).InsertTable(chunkTable);
                    }
                }
            }

            wb.SaveAs(path);
        }
    }
}
