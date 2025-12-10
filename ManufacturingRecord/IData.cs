using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ManufacturingRecord.Data
{
    internal interface IData
    {
        string SearchSqlFile();
        void QueryMachineManufacturingResume(string product, string feature, string process, DateTime fromDate, DateTime toDate, DataGridView dgv);
    }
}
