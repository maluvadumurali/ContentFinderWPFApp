using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
namespace ContentFinderWPFApp
{
    internal class WorkBookData
    {
        public bool IsValid { get; set; }
        public Excel.Application ExcelApplication { get; set; }
        public Excel.Workbook ExcelWorkbook { get; set; }
        public Excel._Worksheet ExcelSheet { get; set; }
        public Excel.Range ExcelRange { get; set; }
        public int RowCount { get; set; }
        public Dictionary<string, string> ExcelData { get; set; }
    }
}