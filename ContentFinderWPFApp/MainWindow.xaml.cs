using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Fiddler;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ContentFinderWPFApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, string> excelData = new Dictionary<string, string>();

        bool isApplicationRunning = false;
        bool isExceptionInReadingData = false;
        string filePath;
        int columnNumber, sheetNumber, idColumn;

        IList<string> matchingIds = new List<string>();

        delegate void GetUrl();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ApplicationAfterSessionComplete(Session oSession)
        {
            listBoxURL.Dispatcher.Invoke(new GetUrl(() =>
            {
                listBoxURL.Items.Add(oSession.url);
                string respHeaders = oSession.oResponse.headers.ToString();
                string respBody = oSession.GetResponseBodyAsString();
                MatchWithExcel(respBody);

            }));
        }

        private void MatchWithExcel(string body)
        {
            if (!excelData.Any())
            {
                if (!isExceptionInReadingData)
                {
                    excelData = GetExcelData(filePath, columnNumber, sheetNumber, idColumn);
                }
                else
                {
                    listBoxLoadedIds.BorderThickness = new Thickness(12);
                    listBoxLoadedIds.BorderBrush = Brushes.Red;
                    listBoxLoadedIds.Items.Add("Something went wrong while reading data from Excel.");
                    listBoxLoadedIds.Items.Add("Please verify inputs and click run button again.");
                    CloseFiddler();
                }
            }
            if (excelData.Any())
            {
                listBoxFound.BorderBrush = Brushes.Green;
                if (!String.IsNullOrEmpty(body))
                {
                    foreach (var data in excelData)
                    {
                        if (body.Contains(data.Value))
                        //if (Regex.IsMatch(body.Trim(' '), data.Trim(' '), RegexOptions.IgnoreCase))
                        {
                            if (!string.IsNullOrEmpty(data.Value))
                            {
                                listBoxFound.Items.Add(data.Key);
                                matchingIds.Add(data.Key);
                            }
                        }

                    }
                }
            }
            else
            {
                listBoxFound.Items.Add("Please Check Excel Sheet whether it has specified sheet number and column number or not");
                listBoxFound.BorderBrush = Brushes.Red;
                listBoxFound.BorderThickness = new Thickness(12);
            }
        }

        private Dictionary<string, string> GetExcelData(string filePath, int columnNumber, int sheetNumber, int idColumn)
        {
            WorkBookData workBook = GetExcelDocument(filePath, sheetNumber, columnNumber, idColumn);
            Dictionary<string, string> returnData = new Dictionary<string, string>();

            if (workBook.IsValid)
            {
                returnData = workBook.ExcelData;
                listBoxFound.Items.Clear();
                txtPath.BorderBrush = Brushes.Green;
                txtColumn.BorderBrush = Brushes.Green;
                txtIds.BorderBrush = Brushes.Green;
                txtSheet.BorderBrush = Brushes.Green;
            }
            else
            {
                txtPath.BorderBrush = Brushes.Red;
                txtColumn.BorderBrush = Brushes.Red;
                txtIds.BorderBrush = Brushes.Red;
                txtSheet.BorderBrush = Brushes.Red;
            }

            return returnData;
        }

        private WorkBookData GetExcelDocument(string filePath, int sheetNumber, int columnNumber, int idColumnNumber)
        {
            WorkBookData data = new WorkBookData { IsValid = false };
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(@filePath);
                xlWorksheet = xlWorkbook.Sheets[sheetNumber];
                xlRange = xlWorksheet.UsedRange;

                data.ExcelApplication = xlApp;
                data.ExcelRange = xlRange;
                data.ExcelSheet = xlWorksheet;
                data.ExcelWorkbook = xlWorkbook;
                data.ExcelData = new Dictionary<string, string>();
                var rowCount = xlRange.Rows.Count;
                for (int i = 0; i < rowCount; i++)
                {
                    if (xlRange.Cells[i + 1, idColumnNumber].Value2 != null && xlRange.Cells[i + 1, columnNumber].Value2 != null)
                    {
                        data.ExcelData.Add(xlRange.Cells[i + 1, idColumnNumber].Value2.ToString(), xlRange.Cells[i + 1, columnNumber].Value2.ToString());
                        listBoxLoadedIds.Items.Add(xlRange.Cells[i + 1, idColumnNumber].Value2.ToString() + ": " + xlRange.Cells[i + 1, columnNumber].Value2.ToString());
                    }

                }
                if (data.ExcelData.Count > 0)
                {
                    data.IsValid = true;
                }
                else
                {
                    isExceptionInReadingData = true;
                }
            }
            catch (Exception ex)
            {
                isExceptionInReadingData = true;
                listBoxLoadedIds.Items.Add(ex.Message);
                listBoxLoadedIds.BorderBrush = Brushes.Red;
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad
                //release com objects to fully kill excel process from running in the background
                if (xlRange != null)
                {
                    Marshal.ReleaseComObject(xlRange);
                }
                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }

                

                if(xlWorkbook != null)
                {
                    //close and release
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            return data;
        }

        private void main_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (isApplicationRunning)
            {
                CloseFiddler();
                if (SaveReportToExcel(excelData))
                {
                    UninstallCertificate();
                }
            }
        }

        private static void CloseFiddler()
        {
            if (FiddlerApplication.IsStarted())
            {
                FiddlerApplication.oProxy.Detach();
                Fiddler.FiddlerApplication.Shutdown();
            }
        }

        private bool SaveReportToExcel(Dictionary<string, string> excelData)
        {
            Dictionary<string, string> reportData = new Dictionary<string, string>();

            foreach (var data in excelData)
            {
                if (matchingIds.Contains(data.Key))
                {
                    reportData.Add(data.Key, "Found");
                }
                else
                {
                    reportData.Add(data.Key, "Not Found");
                }
            }

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Found/ Not found";

            xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 2]].Interior.Color = System.Drawing.Color.Blue;

            int row = 1;
            foreach (var data in reportData)
            {
                
                xlWorkSheet.Cells[row + 1, 1] = data.Key;
                xlWorkSheet.Cells[row + 1, 2] = data.Value;
                if(data.Value == "Not Found")
                {
                    xlWorkSheet.Range[xlWorkSheet.Cells[row + 1, 1], xlWorkSheet.Cells[row+1, 2]].Interior.Color = System.Drawing.Color.Red;
                }
                else
                {
                    xlWorkSheet.Range[xlWorkSheet.Cells[row + 1, 1], xlWorkSheet.Cells[row + 1, 2]].Interior.Color = System.Drawing.Color.Green;
                }
                row++;
            }

            string trimmmedPath = txtPath.Text.Trim('\\');
            string[] xlPath = trimmmedPath.Split('\\');
            string fileLocation = string.Empty;
            for (int i = 0; i < xlPath.Length-1; i++)
            {
                fileLocation = fileLocation + xlPath[i] + "\\";
            }

            try
            {
                xlWorkBook.SaveAs(fileLocation + "Report.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch
            {
                // Do nothing
            }
            finally
            {
                xlWorkBook.Close(false, misValue, misValue);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            return true;
        }

        private void txtPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearExcelData();
        }

        private void txtSheet_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearExcelData();
        }

        private void txtColumn_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearExcelData();
        }

        private void ClearExcelData()
        {
            excelData = new Dictionary<string, string>();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            // Set filter for file extension and default file extension

            //dlg.DefaultExt = ".xlsx | .xls";
            dlg.Filter = "Excel documents (.xlsx, .xls)|*.xlsx; *.xls";

            // Display OpenFileDialog by calling ShowDialog method

            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox

            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;

                txtPath.Text = filename;
                ClearExcelData();
            }

        }

        private void Button_LoadAndRun_Click(object sender, RoutedEventArgs e)
        {
            listBoxLoadedIds.BorderBrush = Brushes.Green;
            listBoxLoadedIds.Items.Clear();

            filePath = txtPath.Text.Trim();
            columnNumber = Convert.ToInt32(txtColumn.Text.Trim());
            sheetNumber = Convert.ToInt32(txtSheet.Text.Trim());
            idColumn = Convert.ToInt32(txtIds.Text.Trim());
            excelData = GetExcelData(filePath, columnNumber, sheetNumber, idColumn);

            InstallCertificate();

            Fiddler.FiddlerApplication.AfterSessionComplete += ApplicationAfterSessionComplete;
            if (!Fiddler.FiddlerApplication.IsStarted())
            {
                Fiddler.FiddlerApplication.Startup(0, true, true, true);
            }

            this.isApplicationRunning = true;
        }

        public static bool InstallCertificate()
        {
            if (!CertMaker.rootCertExists())
            {
                if (!CertMaker.createRootCert())
                    return false;

                if (!CertMaker.trustRootCert())
                    return false;
            }

            return true;
        }

        public static bool UninstallCertificate()
        {
            if (CertMaker.rootCertExists())
            {
                if (!CertMaker.removeFiddlerGeneratedCerts(true))
                    return false;
            }
            return true;
        }

    }
}
