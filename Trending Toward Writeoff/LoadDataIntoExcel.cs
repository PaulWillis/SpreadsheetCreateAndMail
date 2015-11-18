using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = NetOffice.ExcelApi;
using System.Configuration;
using System.Globalization;
using System.Collections.Specialized;

namespace Trending_Toward_Writeoff
{
    class DataSetWithName
    {
        public string DatasetName { get; set; }
        public DataSet DataSet{ get; set; }

    }
    class LoadDataIntoExcel
    { 

        private static NameValueCollection appConfig = ConfigurationManager.AppSettings;

        public void Run()
        {

            Console.WriteLine("running");
             
            Console.WriteLine(DateTime.Now.ToString()); 
            if (checkRequiredFields())
            { 
                List<DataSetWithName> listOfDataSets = new List<DataSetWithName>();
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("21"), DatasetName = "All" });
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("21"), DatasetName = "21" });
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("10"), DatasetName = "10" });
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("7"), DatasetName = "7" });
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("3"), DatasetName = "3" });
                listOfDataSets.Add(new DataSetWithName { DataSet = DAL.Instance.GetDataset("1"), DatasetName = "1" });
                               
                
                string FileSavedAs = "";
                CreateNewExcelDoc(listOfDataSets, appConfig["SpreadsheetFileName"], out FileSavedAs);
            

                MailThis(appConfig["EmailAddress"], FileSavedAs);
            }
             
        }

        static bool checkRequiredFields()
        {
            if (string.IsNullOrEmpty(appConfig["EmailAddress"]))
            {
                Console.WriteLine("EmailAddress was not set in the App.config file.");
                return false;
            }

            if (string.IsNullOrEmpty(appConfig["SpreadsheetFileName"]))
            {
                Console.WriteLine("SpreadsheetFileName was not set in the App.config file.");
                return false;
            }

            return true;

        }

        private void MailThis(string ToCSV, string FileSafedAs)
        {
            Mailer m = new Mailer(ToCSV
                                    , "Monthly - Trending Toward Write Offs "
                                    , @"</br>Attached is a report aging accounts.
                                        </br>
                                        </br>
                                        </br>
                                        </br>Executed on "
                                    , FileSafedAs);

        }


        public void AddDatasetIntoWorksheet(string Title,string workSheetName, DataSet ds, Excel.Worksheet workSheet, Excel.Workbook workbook)
        {

            // font action
            workSheet.Name = workSheetName;

            workSheet.Range("A2").Value = Title;
            workSheet.Range("A2").Font.Name = "Arial";
            workSheet.Range("A2").Font.Size = 10;
            workSheet.Range("A2").Font.Bold = true;
            workSheet.Range("A2").Font.Italic = true;
            workSheet.Range("A2").Font.Underline = true;
             
            // setup rows and columns
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].ColumnWidth = 25;
            workSheet.Columns[3].ColumnWidth = 25;

            workSheet.Columns[4].ColumnWidth = 25;
            workSheet.Columns[5].ColumnWidth = 25; 
            CultureInfo cultureInfo = NetOffice.Settings.ThreadCulture;

            List<string> colNamesToHighlight = GetColNamesToHighlight();
            List<string> colNamesToHighlight_Red = GetColNamesToHighlight_Red();

            int RowToStartHighlight = 0;
            int TotalsRow = 0;
            decimal TotalOfInactiveOver90 = 0;
            decimal TotalOfInactiveOver60 = 0;
            List<string> alpha = AToZ();
            int rowNum = 5;
            int colcount = 0;

            //set certain columns to text format
            workSheet.Columns[3].NumberFormat = "@";
            workSheet.Columns[14].NumberFormat = "@";
            workSheet.Columns[17].NumberFormat = "@";
            workSheet.Columns[18].NumberFormat = "@";



            try
            { 

                foreach (DataTable table in ds.Tables)
                {
                    //write column headings

                    foreach (DataRow row in table.Rows)
                    {
                        foreach (DataColumn column in table.Columns)
                        {

                            workSheet.Range(alpha[colcount] + rowNum.ToString()).Value = column.ColumnName;
                            workSheet.Range(alpha[colcount] + rowNum.ToString()).Font.Bold = true; 
                            colcount += 1;
                        } 
                        break;
                    }

                    colcount = 0;
                    rowNum = 6; 
                    foreach (DataRow row in table.Rows)
                    {
                        foreach (DataColumn column in table.Columns)
                        {
                             
                            workSheet.Range(alpha[colcount] + rowNum.ToString()).Value = row[column];
                            string Pattern2 = string.Format("#{1}##0{0}00",
                            cultureInfo.NumberFormat.CurrencyDecimalSeparator,
                            cultureInfo.NumberFormat.CurrencyGroupSeparator);
                            workSheet.Range(alpha[colcount] + rowNum.ToString()).NumberFormat = Pattern2; 

                            if (row[column].ToString().ToLower() == "inactive>60")
                                RowToStartHighlight = rowNum;

                            if (row[column].ToString().ToLower() == "totals")
                                TotalsRow = rowNum;

                            if (TotalsRow == rowNum)
                            {
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Font.Bold = true;
                            }

                            if (IsColToHighlight(colNamesToHighlight, column.ColumnName) && RowToStartHighlight == rowNum)
                            { 
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Font.Color = ToDouble(System.Drawing.Color.Green);
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Interior.Color = ToDouble(System.Drawing.Color.FromArgb(198, 239, 206));
                                TotalOfInactiveOver90 += Convert.ToDecimal(workSheet.Range(alpha[colcount] + rowNum.ToString()).Value); 
                            }

                            if (IsColToHighlight(colNamesToHighlight_Red, column.ColumnName) && RowToStartHighlight == rowNum)
                            { 
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Font.Color = ToDouble(System.Drawing.Color.DarkRed);
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Interior.Color = ToDouble(System.Drawing.Color.FromArgb(198, 239, 206));
                                workSheet.Range(alpha[colcount] + rowNum.ToString()).Interior.Color = ToDouble(System.Drawing.Color.FromArgb(204, 128, 51));
                                TotalOfInactiveOver60 += Convert.ToDecimal(workSheet.Range(alpha[colcount] + rowNum.ToString()).Value);
                            }

                            colcount += 1;
                        }
                        rowNum += 1;
                        colcount = 0;
                    }
                }

                int newR = rowNum + 1;

                workSheet.Columns[1].ColumnWidth = 20;
                workSheet.Columns[2].ColumnWidth = 12;
                workSheet.Columns[3].ColumnWidth = 12;
                workSheet.Columns[4].ColumnWidth = 12;
                workSheet.Columns[5].ColumnWidth = 12;
                workSheet.Columns[6].ColumnWidth = 12;
                workSheet.Columns[7].ColumnWidth = 12;
                workSheet.Columns[8].ColumnWidth = 12;
                workSheet.Columns[9].ColumnWidth = 12;
                workSheet.Columns[10].ColumnWidth = 12;
                workSheet.Columns[11].ColumnWidth = 12;
                workSheet.Columns[12].ColumnWidth = 12;
                workSheet.Columns[13].ColumnWidth = 12;
                
            }
            catch (Exception ex)
            {
                workSheet.Range("A100").Value = ex.Message;
            }
        }



        public void CreateNewExcelDoc(List<DataSetWithName> listOfDataSets, string SpreadsheetFileName, out string FileSavedAs)
        {

            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();


            int wksheetCount = 1;
            foreach (DataSetWithName ds in listOfDataSets)
            {
                AddDatasetIntoWorksheet("Cycle: " + ds.DatasetName.ToString(), "Cycle" + ds.DatasetName, ds.DataSet,(Excel.Worksheet)workBook.Worksheets.Add(), workBook);
                wksheetCount += 1;
            }


            // save it 
            string fileExtension = GetDefaultExtension(excelApplication);
            string WorkingDirectory = Environment.CurrentDirectory; 

            string workbookFile = string.Format(WorkingDirectory + "\\" + SpreadsheetFileName + "{0}", fileExtension);
            DeleteIfExists(workbookFile);

            workBook.SaveAs(workbookFile);

            FileSavedAs = workbookFile;


            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();
             


        }

        private void DeleteIfExists(string workbookFile)
        {
            if (System.IO.File.Exists(workbookFile))
                System.IO.File.Delete(workbookFile);
        }

        private static double ToDouble(System.Drawing.Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }


        private bool IsColToHighlight(List<string> ColNamesToHighlight, string ColumnName)
        { 
            if (ColNamesToHighlight.Any(x => x == ColumnName))
                return true;

            return false;
             
        }

        private List<string> GetColNamesToHighlight_Red()
        {
            List<string> res = new List<string>();
            res.Add("60 Days");
            return res;
        }

        private List<string> GetColNamesToHighlight()
        {
            List<string> res = new List<string>();
            res.Add("90 Days");
            res.Add("120 Days");
            res.Add("150 Days");
            res.Add("180 Days");
            return res;
        }

        private List<string> AToZ()
        {
            List<string> res = new List<string>();

            for (int i = 0; i < 26; i++)
            {
                string p = Convert.ToChar((Convert.ToInt32('A') + i)).ToString();
                res.Add(p);
            }
            return res;
        }


        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".xlsx";
            else
                return ".xls";
        }

        private void CreateNewExcelDoc0(DataSet ds, string FileName)
        {
            System.Data.DataTable tbl = ds.Tables[0];

            System.IO.StreamWriter file = new System.IO.StreamWriter(FileName);

            string cleanline = "";
            foreach (DataColumn c in tbl.Columns)
            {
                cleanline += c.ColumnName + ",";
            }
            cleanline = cleanline.Substring(0, cleanline.Length - 1);
            file.Write(cleanline);
            file.Write(Environment.NewLine);
            foreach (DataRow r in tbl.Rows)
            {
                cleanline = "";
                for (int i = 0; i < tbl.Columns.Count; i++)
                {
                    cleanline += r[i].ToString().Replace(",", string.Empty) + ","; 
                }
                cleanline = cleanline.Substring(0, cleanline.Length - 1);
                file.Write(cleanline);
                file.Write(Environment.NewLine);
            }
            file.Close();


        }
    }
}