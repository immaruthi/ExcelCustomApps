using ExcelDataReader;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AppServicesForExcel
{
    public partial class AppServiceForExcel : ServiceBase
    {

        private Timer Schedular;
        int dueTime;
        private volatile bool _requestStop = false;


        public AppServiceForExcel()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            dueTime = int.Parse(System.Configuration.ConfigurationSettings.AppSettings["IntervalMinutes"]) * 60000;
            Schedular = new Timer(new TimerCallback(ScheduleTasksCallBack));
            ScheduleTasksCallBack(null);
        }

        private void SecondWay()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            string filePath = @"C:\Kalya Solutions\GMB insights.xlsx";

            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + filePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";
            DataTable resultsDataset = new DataTable();
            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [GMB insights (Discovery Report)$]", conn);
                conn.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(resultsDataset);
            }

            if (resultsDataset.Rows.Count > 3)
            {

                //add data 
                xlWorkSheet.Cells[1, 1] = "";
                xlWorkSheet.Cells[1, 2] = "Student1";
                xlWorkSheet.Cells[1, 3] = "Student2";
                xlWorkSheet.Cells[1, 4] = "Student3";

                xlWorkSheet.Cells[2, 1] = "Term1";
                xlWorkSheet.Cells[2, 2] = "80";
                xlWorkSheet.Cells[2, 3] = "65";
                xlWorkSheet.Cells[2, 4] = "45";

                xlWorkSheet.Cells[3, 1] = "Term2";
                xlWorkSheet.Cells[3, 2] = "78";
                xlWorkSheet.Cells[3, 3] = "72";
                xlWorkSheet.Cells[3, 4] = "60";

                xlWorkSheet.Cells[4, 1] = "Term3";
                xlWorkSheet.Cells[4, 2] = "82";
                xlWorkSheet.Cells[4, 3] = "80";
                xlWorkSheet.Cells[4, 4] = "65";

                xlWorkSheet.Cells[5, 1] = "Term4";
                xlWorkSheet.Cells[5, 2] = "75";
                xlWorkSheet.Cells[5, 3] = "82";
                xlWorkSheet.Cells[5, 4] = "68";

                xlWorkSheet.Cells[6, 1] = "Term6";
                xlWorkSheet.Cells[6, 2] = "75";
                xlWorkSheet.Cells[6, 3] = "82";
                xlWorkSheet.Cells[6, 4] = "68";

                xlWorkSheet.Cells[7, 1] = "Term7";
                xlWorkSheet.Cells[7, 2] = "75";
                xlWorkSheet.Cells[7, 3] = "82";
                xlWorkSheet.Cells[7, 4] = "68";

                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 200, 300, 250);
                Excel.Chart chartPage = myChart.Chart;
                //chartPage.Location(Excel.XlChartLocation.xlLocationAutomatic, "Chart1");

                chartRange = xlWorkSheet.get_Range("A1", "d5");
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlDoughnutExploded;

                xlWorkBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, "Test3.pdf");
            }

            //xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            //MessageBox.Show("Excel file created , you can find the file c:\\csharp.net-informations.xls");

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void DirectoryCreation(string[] recommendedInputPaths)
        {
            //<add key="CentralizedDB" value="Data Source=maruthi-t470\sqlexpress;Initial Catalog=DB_UPS_ADTRS;Integrated Security=True"/>
            try
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    foreach (string recommendedPath in recommendedInputPaths)
                    {
                        if (!System.IO.Directory.Exists(recommendedPath))
                        {
                            eventLog.WriteEntry("Directories Initialization " + recommendedPath);
                            System.IO.Directory.CreateDirectory(recommendedPath);
                            eventLog.WriteEntry("Directories Initialization " + recommendedPath);
                        }
                    }
                }
            }
            catch (Exception exeption)
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(exeption.Message, EventLogEntryType.Error);
                }
            }
        }

        private void ThrirdWay()
        {

            string[] recommendedInputPaths = new string[] { @"C:\Automation\Input", @"C:\Automation\OutPut", @"C:\Automation\Processed" };


            using (EventLog eventLog = new EventLog())
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry("Files Created", EventLogEntryType.Error);
            }

            try
            {

                
                DirectoryCreation(recommendedInputPaths);

                //string[] accdbDirectory = System.IO.Directory.GetFiles(recommendedInputPaths[0], "*.xlsx");

                //string filePath = @"C:\Kalya Solutions\GMB insights (Discovery Report) - 2020-1-12 - 2020-1-18 - 54dfdd99d13944a11e53e473d2e7a7b4.xlsx";

                

                    foreach (string accdbFile in Directory.GetFiles(recommendedInputPaths[0], "*.xlsx"))
                    {

                    using (EventLog eventLog = new EventLog())
                    {
                        eventLog.Source = "Application";
                        eventLog.WriteEntry("Files Read", EventLogEntryType.SuccessAudit);
                    }

                    //string connectioFile = accdbFile;

                    // string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + accdbFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";
                    DataTable resultsDataset = new DataTable();
                    //    using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + accdbFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";"))
                    //    {
                    //        OleDbCommand cmd = new OleDbCommand("SELECT * FROM [GMB insights (Discovery Report)$]", conn);
                    //        conn.Open();
                    //        OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                    //        adapter.Fill(resultsDataset);
                    //        conn.Dispose();

                    //        resultsDataset.Dispose();
                    //    }


                        IExcelDataReader excelReader;
                        FileStream stream = File.Open(accdbFile, FileMode.Open, FileAccess.Read); 
                        excelReader = ExcelReaderFactory.CreateReader(stream);
                       
                        DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        excelReader.Close();

                    resultsDataset = result.Tables[0];

                    using (EventLog eventLog = new EventLog())
                    {
                        eventLog.Source = "Application";
                        eventLog.WriteEntry("Data Set Prepared", EventLogEntryType.SuccessAudit);
                    }

                    //string moveDestinationPath = Path.Combine(recommendedInputPaths[2], Path.GetFileNameWithoutExtension("GMB insights (Discovery Report) - 2020-1-12 - 2020-1-18 - 54dfdd99d13944a11e53e473d2e7a7b4.xlsx"));

                    //File.Move(@"C:\Automation\Input\GMB insights (Discovery Report) - 2020-1-12 - 2020-1-18 - 54dfdd99d13944a11e53e473d2e7a7b4.xlsx", moveDestinationPath + DateTime.Now.ToString("_MMMdd_yyyy_HHmmss") + ".xlsx");

                    int resultDataSetRow = 1;


                        for (int i = 0; i < 6; i++)
                        {
                            DataSet ds = new DataSet();
                            DataTable mergeTable = new DataTable();
                            mergeTable.Clear();

                            mergeTable.Columns.Add("Storecode", typeof(System.String));
                            mergeTable.Columns.Add("BusinessName", typeof(System.String));
                            mergeTable.Columns.Add("Address", typeof(System.String));
                            mergeTable.Columns.Add("Labels", typeof(System.String));
                            mergeTable.Columns.Add("OverallRating", typeof(System.Double));
                            mergeTable.Columns.Add("TotalSearches", typeof(System.Double));
                            mergeTable.Columns.Add("DirectSearches", typeof(System.Double));
                            mergeTable.Columns.Add("DiscoverySearches", typeof(System.Double));
                            mergeTable.Columns.Add("TotalViews", typeof(System.Double));
                            mergeTable.Columns.Add("SearchViews", typeof(System.Double));
                            mergeTable.Columns.Add("MapsViews", typeof(System.Double));
                            mergeTable.Columns.Add("TotalActions", typeof(System.Double));
                            mergeTable.Columns.Add("WebsiteActions", typeof(System.Double));
                            mergeTable.Columns.Add("DirectionsActions", typeof(System.Double));
                            mergeTable.Columns.Add("PhoneCallActions", typeof(System.Double));


                            DataRow workRow = mergeTable.NewRow();
                            mergeTable.Rows.Add(workRow);

                            mergeTable.Rows[0][0] = resultsDataset.Rows[resultDataSetRow]["Store code"];
                            mergeTable.Rows[0][1] = resultsDataset.Rows[resultDataSetRow]["Business name"];
                            mergeTable.Rows[0][2] = resultsDataset.Rows[resultDataSetRow]["Address"];
                            mergeTable.Rows[0][3] = resultsDataset.Rows[resultDataSetRow]["Labels"];
                            mergeTable.Rows[0][4] = resultsDataset.Rows[resultDataSetRow]["Overall rating"];
                            mergeTable.Rows[0][5] = resultsDataset.Rows[resultDataSetRow]["Total searches"];
                            mergeTable.Rows[0][6] = resultsDataset.Rows[resultDataSetRow]["Direct searches"];
                            mergeTable.Rows[0][7] = resultsDataset.Rows[resultDataSetRow]["Discovery searches"];
                            mergeTable.Rows[0][8] = resultsDataset.Rows[resultDataSetRow]["Total views"];
                            mergeTable.Rows[0][9] = resultsDataset.Rows[resultDataSetRow]["Search views"];
                            mergeTable.Rows[0][10] = resultsDataset.Rows[resultDataSetRow]["Maps views"];
                            mergeTable.Rows[0][11] = resultsDataset.Rows[resultDataSetRow]["Total actions"];
                            mergeTable.Rows[0][12] = resultsDataset.Rows[resultDataSetRow]["Website actions"];
                            mergeTable.Rows[0][13] = resultsDataset.Rows[resultDataSetRow]["Directions actions"];
                            mergeTable.Rows[0][14] = resultsDataset.Rows[resultDataSetRow]["Phone call actions"];

                            ds.Tables.Add(mergeTable);

                            //ReportDataModels.GMBInformationDataTable gMBInformationRows = ds;

                            Microsoft.Reporting.WebForms.ReportViewer ReportViewer1 = new Microsoft.Reporting.WebForms.ReportViewer();

                            ReportViewer1.ProcessingMode = Microsoft.Reporting.WebForms.ProcessingMode.Local;

                            //Path.Combine(Directory.GetCurrentDirectory(), @"DBObjs\RawDataReport.txt")

                            ReportViewer1.LocalReport.ReportPath = @"C:\DRL RDLC\Report1.rdlc";
                            Microsoft.Reporting.WebForms.ReportDataSource datasource = new Microsoft.Reporting.WebForms.ReportDataSource("DataSet1", ds.Tables[0]);
                            ReportViewer1.LocalReport.DataSources.Clear();
                            ReportViewer1.LocalReport.DataSources.Add(datasource);

                        using (EventLog eventLog = new EventLog())
                        {
                            eventLog.Source = "Application";
                            eventLog.WriteEntry("Report Created", EventLogEntryType.SuccessAudit);
                        }

                        Warning[] warnings;
                            string[] streamids;
                            string mimeType;
                            string encoding;
                            string filenameExtension;

                            byte[] bytes = ReportViewer1.LocalReport.Render(
                                "PDF", null, out mimeType, out encoding, out filenameExtension,
                                out streamids, out warnings);


                           // string outputFileName = Path.Combine(recommendedInputPaths[1], resultsDataset.Rows[resultDataSetRow]["Business name"] + ".pdf");

                            //@"C:\Users\v-mapall\source\repos\ExcelCustomApps\GMBRDLC Application\" + resultsDataset.Rows[resultDataSetRow]["Business name"] + ".pdf"

                            using (FileStream fs = new FileStream(Path.Combine(recommendedInputPaths[1], resultsDataset.Rows[resultDataSetRow]["Business name"] + ".pdf"), FileMode.Create))
                            {
                                fs.Write(bytes, 0, bytes.Length);
                            }

                        using (EventLog eventLog = new EventLog())
                        {
                            eventLog.Source = "Application";
                            eventLog.WriteEntry("PDF Files Created", EventLogEntryType.SuccessAudit);
                        }

                        resultDataSetRow++;
                        }

                    //string accdbFileName = Path.GetFileName(accdbFile);

                    //string moveDestinationPath = Path.Combine(recommendedInputPaths[2], Path.GetFileName(accdbFile));


                    string moveDestinationPath = Path.Combine(recommendedInputPaths[2], Path.GetFileNameWithoutExtension(accdbFile));

                    File.Move(accdbFile, moveDestinationPath + DateTime.Now.ToString("_MMMdd_yyyy_HHmmss") + ".xlsx");

                    using (EventLog eventLog = new EventLog())
                    {
                        eventLog.Source = "Application";
                        eventLog.WriteEntry("Files Were Moved", EventLogEntryType.SuccessAudit);
                    }

                }
            }
            catch(Exception exeption)
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(exeption.Message, EventLogEntryType.Error);
                }
            }

        }



        private void ScheduleTasksCallBack(object e)
        {
            try
            {

                if (_requestStop)
                {
                    return;
                }

                ThrirdWay();

                //string[] recommendedInputPaths = new string[] { @"C:\Automation\Input", @"C:\Automation\OutPut", @"C:\Automation\Processed" };

                

                //ExcelGraphGeneration();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error" + ex.Message.ToString());
            }

            Schedular.Change(dueTime, Timeout.Infinite);
        }

        private static void ExcelGraphGeneration()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook
               .Worksheets.get_Item(1);

            // Add data columns
            xlWorkSheet.Cells[1, 1] = "SL";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[1, 3] = "CTC";
            xlWorkSheet.Cells[1, 4] = "DA";
            xlWorkSheet.Cells[1, 5] = "HRA";
            xlWorkSheet.Cells[1, 6] = "Conveyance";
            xlWorkSheet.Cells[1, 7] = "Medical Expenses";
            xlWorkSheet.Cells[1, 8] = "Special";
            xlWorkSheet.Cells[1, 9] = "Bonus";
            xlWorkSheet.Cells[1, 10] = "TA";
            xlWorkSheet.Cells[1, 11] = "TOTAL";
            xlWorkSheet.Cells[1, 11] = "Contribution to PF";
            xlWorkSheet.Cells[1, 12] = "Profession Tax";
            xlWorkSheet.Cells[1, 13] = "TDS";
            xlWorkSheet.Cells[1, 14] = "Salary Advance";
            xlWorkSheet.Cells[1, 15] = "TOTAL";
            xlWorkSheet.Cells[1, 16] = "NET PAY";


            Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp1.Workbooks.Open
               (@"C:\Users\v-mapall\source\repos\ExcelCustomApps\AppServicesForExcel\Sample Data\Sample Data2.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //for (int i = 1; i <= rowCount; i++)

            for (int i = 1; i <= 2; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    Console.WriteLine(xlRange.Cells[i, j].Value2.ToString());
                    xlWorkSheet.Cells[i, j] = xlRange.Cells[i, j]
                       .Value2.ToString();

                }
            }

            //Console.ReadLine();

            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)
               xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)
               xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "R22");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

            // Export chart as picture file
            chartPage.Export(@"C:\Users\v-mapall\source\repos\ExcelCustomApps\AppServicesForExcel\Sample Data\EmployeeExportData.pdf",
               "PDF", misValue);

            xlWorkBook.SaveAs("EmployeeExportData.xls",
               Excel.XlFileFormat.xlWorkbookNormal, misValue,
               misValue, misValue, misValue,
               Excel.XlSaveAsAccessMode.xlExclusive, misValue,
               misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            DeallocateObject(xlWorkSheet);
            DeallocateObject(xlWorkBook);
            DeallocateObject(xlApp);
            DeallocateObject(xlApp1);
        }

        private static void DeallocateObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal
                   .ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occurred while releasingobject " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        protected override void OnStop()
        {
            _requestStop = true;
        }
    }
}
