using ManufacturingRecord.Data;
using ManufacturingRecord.Service;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Windows.Forms;
// ... (引入 service 層的命名空間) ...

namespace ManufacturingRecord
{
    // 🌟 必須使用 partial 關鍵字
    public partial class Form1
    {
        // 🌟 新增一個方法來註冊所有事件
        private void AddEventHandlers()
        {
            _searchButton.Click += _searchButton_Click;
            _exportExcelButton.Click += _exportExcelButton_Click;
            // ... 其他事件 (例如 DGV 的 CellClick)
        }

        private void _searchButton_Click(object sender, EventArgs e)
        {
            // 1. 取得使用者輸入的參數值
            string inputProduct = _productTextBox.Text.Trim().ToString();
            string inputFeature = _featureTextBox.Text.Trim().ToString();
            string inputProcess = _processTextBox.Text.Trim().ToString();
            DateTime inputFromDate = _fromDateTimePicker.Value.Date;
            DateTime inputToDate = _toDateTimePicker.Value.Date;

            if (string.IsNullOrEmpty(inputProduct) || string.IsNullOrEmpty(inputFeature)
                || string.IsNullOrEmpty(inputProcess) || string.IsNullOrEmpty(inputFromDate.ToString()) || string.IsNullOrEmpty(inputToDate.ToString()))
            {
                MessageBox.Show("請確保欄位輸入完整，不能有空值。", "輸入錯誤");
                return;
            }

            var db = new ManufacturingRecord.Data.Data();
            db.QueryMachineManufacturingResume(inputProduct, inputFeature, inputProcess, inputFromDate, inputToDate, dgv);
        }

        private void _exportExcelButton_Click(object sender, EventArgs e)
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
                FileName = "BomRouting_View.xlsx"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                //  直接匯出當前的 DataGridView 狀態
                //_excelService.ExportGridToExcel(dgv, sfd.FileName);
                MessageBox.Show("匯出完成！（以目前 Grid 顯示為準）");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
        }
    }
}