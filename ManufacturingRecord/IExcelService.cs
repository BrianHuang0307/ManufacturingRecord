using System.Data;
using System.Windows.Forms;

namespace ManufacturingRecord.Service
{
    public interface IExcelService
    {
        /// 直接把當前 DataGridView 的畫面匯出到 Excel。
        void ExportGridToExcel(DataGridView grid, string path);
    }
}
