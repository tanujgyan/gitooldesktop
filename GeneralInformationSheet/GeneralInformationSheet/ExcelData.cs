using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace GeneralInformationSheet
{
    public class ExcelData
    {

        public static ArrayList ExcelDataValues = new ArrayList();
        public static List<ArrayList> ExcelDataValuesCumulative = new List<ArrayList>();
        public DataView Data
        {
           
                get
            {

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook;
                    Microsoft.Office.Interop.Excel.Worksheet worksheet;
                    Microsoft.Office.Interop.Excel.Range range;
                    workbook = excelApp.Workbooks.Open(MainWindow.inputFilePath);
                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet3"];
                    MainWindow.worker.ReportProgress(10);
                    int column = 0;
                    int row = 0;

                    range = worksheet.UsedRange;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Last Name");
                    dt.Columns.Add("First Name");
                    dt.Columns.Add("Middle Name");
                    dt.Columns.Add("Banner ID");
                    dt.Columns.Add("Visa Type");
                    dt.Columns.Add("SEVIS ID");
                    dt.Columns.Add("Date Of Birth");
                    dt.Columns.Add("Country Of Birth");
                    dt.Columns.Add("Gender");
                    dt.Columns.Add("Country Of Citizenship");
                    dt.Columns.Add("OSU EmailID");
                    dt.Columns.Add("Personal EmailID");
                    dt.Columns.Add("Street Address 1");
                    dt.Columns.Add("Street Address 2");
                    dt.Columns.Add("City");
                    dt.Columns.Add("State");
                    dt.Columns.Add("Country");
                    dt.Columns.Add("Zip Code");
                    dt.Columns.Add("Admit Term");
                    for (row = 2; row <= range.Rows.Count; row++)
                    {
                        DataRow dr = dt.NewRow();
                        ExcelDataValues = new ArrayList();

                        for (column = 1; column <= range.Columns.Count; column++)
                        {

                            dr[column - 1] = Convert.ToString((range.Cells[row, column] as Microsoft.Office.Interop.Excel.Range).Value2);
                            ExcelDataValues.Add(dr[column - 1]);
                        }
                        ExcelDataValuesCumulative.Add(ExcelDataValues);
                        MainWindow.worker.ReportProgress(60);
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();
                    }
                    MainWindow.worker.ReportProgress(100);
                    workbook.Close(true, Missing.Value, Missing.Value);
                    excelApp.Quit();
                    return dt.DefaultView;
                }
            }
           
        }
    }
