using ManufacturingRecord.Data;
using ManufacturingRecord.Service;

namespace ManufacturingRecord
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            IData dataService = new ManufacturingRecord.Data.Data();
            IExcelService excelService = new ExcelService();

            Application.Run(new Form1(dataService, excelService));
        }
    }
}