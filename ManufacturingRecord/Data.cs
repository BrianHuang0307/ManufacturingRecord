using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO; // 需要保留，因為 Stream 和 StreamReader 用得到
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ManufacturingRecord.Data
{
    internal class Data // : IData (如果你有介面請保留)
    {
        private const string connectionString = "";

        /// <summary>
        /// 從內嵌資源中讀取 SQL 檔案內容
        /// </summary>
        private string GetSqlFromResource()
        {
            var assembly = Assembly.GetExecutingAssembly();

            // 資源名稱格式通常為： "專案預設命名空間.資料夾名(如果有的話).檔名"
            // SQL 檔直接放在專案根目錄，且專案名稱是 ManufacturingRecord
            string resourceName = "ManufacturingRecord.MachineManufacturingResume.sql";

            using (Stream? stream = assembly.GetManifestResourceStream(resourceName))
            {
                // 防呆：如果找不到資源，列出所有可用的資源名稱以供除錯
                if (stream == null)
                {
                    string[] existingResources = assembly.GetManifestResourceNames();
                    string errorMsg = $"找不到內嵌資源: {resourceName}\n\n系統中現有的資源如下:\n{string.Join("\n", existingResources)}";
                    throw new Exception(errorMsg);
                }

                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void QueryMachineManufacturingResume(DateTime fromDate, DateTime toDate, DataGridView dgv)
        {
            DataTable dt = new DataTable();

            try
            {
                // 改用內嵌資源讀取 SQL
                string sql = GetSqlFromResource();

                using (var conn = new OracleConnection(connectionString))
                using (var cmd = new OracleCommand(sql, conn))
                using (var adapter = new OracleDataAdapter(cmd))
                {
                    conn.Open();

                    // 設定參數 (注意：Oracle 參數順序可能重要，視你的 BindByName 設定而定，建議開啟 BindByName)
                    cmd.BindByName = true;

                    cmd.Parameters.Add(new OracleParameter("from_date", OracleDbType.Date)).Value = fromDate;
                    cmd.Parameters.Add(new OracleParameter("to_date", OracleDbType.Date)).Value = toDate;

                    dt.Clear();
                    adapter.Fill(dt);

                    // 更新 UI
                    dgv.SuspendLayout();
                    dgv.DataSource = null;
                    dgv.Columns.Clear();
                    dgv.AutoGenerateColumns = true;
                    dgv.DataSource = dt;
                    dgv.ResumeLayout();

                    if (dgv.Rows.Count > 0)
                    {
                        MessageBox.Show($"資料載入成功！共 {dgv.Rows.Count} 筆記錄。", "完成");
                    }
                    else
                    {
                        MessageBox.Show("查詢成功，但沒有符合條件的資料。", "無資料");
                    }
                }
            }
            catch (OracleException ex)
            {
                MessageBox.Show($"資料庫查詢發生錯誤：\n{ex.Message}", "資料庫錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // 這裡會捕捉到找不到資源的錯誤 (如果是資源名稱打錯)
                MessageBox.Show($"系統發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}