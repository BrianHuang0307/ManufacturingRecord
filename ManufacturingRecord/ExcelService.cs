using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ManufacturingRecord.Service
{
    internal class ExcelService : IExcelService
    {
        public void ExportGridToExcel(DataGridView grid, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("機器生產履歷");

            int col = 1;
            var exportColumns = grid.Columns
                .Cast<DataGridViewColumn>()
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            foreach (var c in exportColumns)
            {
                ws.Cell(1, col).Value = c.HeaderText;
                col++;
            }

            int rowIndex = 2;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;

                col = 1;
                foreach (var c in exportColumns)
                {
                    var cell = row.Cells[c.Name];
                    var v = cell.Value;
                    ws.Cell(rowIndex, col).Value = v == null ? "" : v.ToString();
                    col++;
                }

                rowIndex++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }
    }
}
