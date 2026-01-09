using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ManufacturingRecord.Data
{
    public interface IData
    {
        DataTable QueryMachineManufacturingResume(DateTime fromDate, DateTime toDate);//, DataGridView dgv, string product, string feature, string process);
    }
}
