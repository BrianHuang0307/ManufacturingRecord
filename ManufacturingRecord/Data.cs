using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ManufacturingRecord.Data
{
    internal class Data : IData
    {
        private const string connectionString = "";
        public const string sqlFileName = "MachineManufacturingResume.sql";

        public string SearchSqlFile()
        {
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            string fullPath = Path.Combine(currentDir, sqlFileName);

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"找不到 SQL 檔案: {fullPath}");
            }

            return File.ReadAllText(fullPath);
        }

        public void QueryMachineManufacturingResume(string product, string feature, string process, DateTime fromDate, DateTime toDate, DataGridView dgv)
        {
            DataTable dt = new DataTable();

            try
            {
                string sql = SearchSqlFile();

                using (var conn = new OracleConnection(connectionString))    
                using (var cmd = new OracleCommand(sql, conn))
                using (var adapter = new OracleDataAdapter(cmd))
                {
                    conn.Open();
                    cmd.Parameters.Add(new OracleParameter("from_date", OracleDbType.Date)).Value = fromDate;
                    cmd.Parameters.Add(new OracleParameter("to_date", OracleDbType.Date)).Value = toDate;
                    cmd.Parameters.Add(new OracleParameter("process", OracleDbType.Varchar2)).Value = process;
                    cmd.Parameters.Add(new OracleParameter("feature", OracleDbType.Varchar2)).Value = feature;
                    cmd.Parameters.Add(new OracleParameter("product", OracleDbType.Varchar2)).Value = product;

                    MessageBox.Show($"{product}\n{feature}\n{process}\n{fromDate}\n{toDate}");
                    dt.Clear();
                    dgv.AutoGenerateColumns = true;
                    adapter.Fill(dt);
                    

                    dgv.DataSource = dt;
                    //dgv.DataSource = dataTable;
                    MessageBox.Show($"資料載入成功！共 {dgv.Rows.Count} 筆記錄。", "完成");
                    dgv.Invalidate();
                    dgv.Refresh();
                }
            }
            catch (OracleException ex)
            {
                MessageBox.Show($"資料庫查詢發生錯誤：\n{ex.Message}");
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"查詢發生錯誤：\n{ex.Message}");
            }
        }
    }
}
