using System.Data;
using System.Windows.Forms;

namespace ManufacturingRecord.Service
{
    public interface IExcelService
    {
        /// 直接把當前 DataGridView 的畫面匯出到 Excel。
        void ExportDataToExcel(DataTable dataTable, string path, string sheetName);
        // void ExportDataToExcel(DataGridView dgv, string path, string sheetName);
    }
}
