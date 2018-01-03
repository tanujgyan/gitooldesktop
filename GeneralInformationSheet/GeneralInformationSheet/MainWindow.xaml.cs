using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
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

namespace GeneralInformationSheet
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public static string inputFilePath = "";
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;
        object misvalue = System.Reflection.Missing.Value;
        BarcodeLib.Barcode b = new BarcodeLib.Barcode();
        public static BackgroundWorker worker;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            // Set filter for file extension and default file extension
            // dlg.DefaultExt = ".xlsx";
            //dlg.Filter = "Text documents (.txt)|*.txt";
            
            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();
            
            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                textBox.Text = filename;
                inputFilePath = filename;
            }
            MessageBox.Show("Data Load Started. We will notify you when it is complete");
           
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.ReportProgress(10);
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.ReportProgress(25);
            worker.ProgressChanged += worker_ProgressChanged;
            worker.DoWork += worker_DoWork;
            worker.ReportProgress(50);
            
            
            worker.RunWorkerAsync();
           
            //ExcelData exceldata = new ExcelData();
            //this.dvPreviewData.DataContext = exceldata;
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            
            for (int j = 0; j <ExcelData.ExcelDataValuesCumulative.Count(); j++)
            {
                try
                {
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;

                    //Get a new workbook.
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                    //set the middle column width as it acts a delimiter column only
                    oSheet.Columns["C:C"].ColumnWidth = 2;
                    oSheet.Columns["A:A"].ColumnWidth = 23;
                    oSheet.Columns["B:B"].ColumnWidth = 23;
                    oSheet.Columns["D:D"].ColumnWidth = 23;
                    oSheet.Columns["E:E"].ColumnWidth = 23;
                    oSheet.Rows.RowHeight = 22.5;
                    oSheet.Rows.WrapText = true;

                    #region RowWidth For Special ROws
                    //Included rows - DOB, Email, Immigration History Table Headers
                    oSheet.Rows[5].RowHeight = 30;
                    oSheet.Rows[7].RowHeight = 30;
                    oSheet.Rows[38].RowHeight = 30;
                    oSheet.Rows[44].RowHeight = 30;
                    #endregion
                    #region Labels Only Include text bold code
                    //Header Code
                    oSheet.Cells[1, 1] = "General Information Sheet";
                  
                    oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 5]].Merge();
                    oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 5]].HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //Last Name and First Name Code - Row 2
                    oSheet.Cells[2, 1] = "Last Name";
                    oSheet.Cells[2, 4] = "First Name";

                    //Middle Name and CWID Code - Row 3
                    oSheet.Cells[3, 1] = "Middle Name";
                    oSheet.Cells[3, 4] = "Banner ID (Student ID)";

                    //Visa Type and SEVIS ID Code - Row 4
                    oSheet.Cells[4, 1] = "Visa Type";
                    oSheet.Cells[4, 4] = "SEVIS ID";

                    //DOB and COB Code - Row 5
                    oSheet.Cells[5, 1] = "Date Of Birth (DD-MMM-YYYY)";
                    oSheet.Cells[5, 4] = "Country Of Birth";

                    //Gender and COC Code - Row 6
                    oSheet.Cells[6, 1] = "Gender";
                    oSheet.Cells[6, 4] = "Country Of Citizenship";

                    //OSU Email Address and Alternate Email Address Code - Row 7
                    oSheet.Cells[7, 1] = "OSU Email Address";
                    oSheet.Cells[7, 4] = "Alternate Email Address";

                    //Row 8 will be empty

                    //Home Country Address Heading Only - Row 9
                    oSheet.Cells[9, 1] = "Home Country Address";
                    oSheet.Range[oSheet.Cells[9, 1], oSheet.Cells[9, 5]].Merge();
                    //Street Address 1 - Row 10
                    oSheet.Cells[10, 1] = "Street Address 1";
                    oSheet.Range[oSheet.Cells[10, 2], oSheet.Cells[10, 5]].Merge();
                    //Street Address 2 - Row 11
                    oSheet.Cells[11, 1] = "Street Address 2";
                    oSheet.Range[oSheet.Cells[11, 2], oSheet.Cells[11, 5]].Merge();

                    //City and State - Row 12
                    oSheet.Cells[12, 1] = "City";
                    oSheet.Cells[12, 4] = "State";

                    //Country and Zip Code - Row 13
                    oSheet.Cells[13, 1] = "Country";
                    oSheet.Cells[13, 4] = "Zip Code";

                    //Row 14 will be empty

                    //Local/Temporary Address Heading Only Row 15
                    oSheet.Cells[15, 1] = "Local/Temporary Address";
                    oSheet.Range[oSheet.Cells[15, 1], oSheet.Cells[15, 5]].Merge();
                    //Street Address 1 - Row 16
                    oSheet.Cells[16, 1] = "Street Address 1";
                    oSheet.Range[oSheet.Cells[16, 2], oSheet.Cells[16, 5]].Merge();
                    //Street Address 2 - Row 17
                    oSheet.Cells[17, 1] = "Street Address 2";
                    oSheet.Range[oSheet.Cells[17, 2], oSheet.Cells[17, 5]].Merge();

                    //City and State - Row 18
                    oSheet.Cells[18, 1] = "City";
                    oSheet.Cells[18, 4] = "State";

                    //Country and Zip Code - Row 19
                    oSheet.Cells[19, 1] = "Country";
                    oSheet.Cells[19, 4] = "Zip Code";

                    //Local Phone - Row 20
                    oSheet.Cells[20, 1] = "Local Phone\n(Leave blank if you don't have a local phone now but update ISS once you do)";
                    oSheet.Range[oSheet.Cells[20, 1], oSheet.Cells[20, 4]].Merge();

                    oSheet.Range[oSheet.Cells[20, 1], oSheet.Cells[20, 4]].RowHeight = 27.75;
                    //Row 21 will be empty

                    //Emergency Contact Information Heading Only Row 22
                    oSheet.Cells[22, 1] = "Emergency Contact Information";
                    oSheet.Range[oSheet.Cells[22, 1], oSheet.Cells[22, 5]].Merge();
                    oSheet.Range[oSheet.Cells[22, 1], oSheet.Cells[22, 5]].HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    //Explanation Text Row 23
                    oSheet.Cells[23, 1] = "Family Member, Friend or relative in the U.S. the ISS office should contact on your behalf in case of medical emergency (such as illness or accident). Please fill as much information as you can and update the ISS office with any changes";
                    oSheet.Range[oSheet.Cells[23, 1], oSheet.Cells[23, 5]].Merge();

                    oSheet.Range[oSheet.Cells[23, 1], oSheet.Cells[23, 5]].RowHeight = 45;


                    //Row 24 will be empty

                    //Name and Relationship Row 25
                    oSheet.Cells[25, 1] = "Name";
                    oSheet.Cells[25, 4] = "Relationship";

                    //Street Address 1 - Row 26
                    oSheet.Cells[26, 1] = "Street Address 1";

                    //Street Address 2 - Row 27
                    oSheet.Cells[27, 1] = "Street Address 2";


                    //City and State - Row 28
                    oSheet.Cells[28, 1] = "City";
                    oSheet.Cells[28, 4] = "State";

                    //Country and Zip Code - Row 29
                    oSheet.Cells[29, 1] = "Zip Code";

                    //Phone and Email Address - Row 30
                    oSheet.Cells[30, 1] = "Phone";
                    oSheet.Cells[30, 4] = "Email Address";

                    //Row 31 is empty
                    //Row 32 is empty
                    //Row 33 is empty
                    //Previous semester field Row 34
                    oSheet.Cells[34, 1] = "If you previously attended OSU enter the last semester you attended(Eg. Fall 2016)";
                    oSheet.Range[oSheet.Cells[34, 1], oSheet.Cells[34, 4]].Merge();
                    //Row 35 is empty

                    //Previously entered US heading Row 36
                    oSheet.Cells[36, 1] = "If you previously entered the U.S. provide information below:";
                    oSheet.Range[oSheet.Cells[36, 1], oSheet.Cells[36, 5]].Merge();

                    //Row 37 is empty

                    //Approximate Date of Entry, Approximate Date of Exit, Immigration Status E.g. F1, F2 etc, Primary Activity E.g. High School, Exchange etc - Row 38
                    oSheet.Cells[38, 1] = "Approximate Date of Entry";
                    oSheet.Cells[38, 2] = "Approximate Date of Exit";
                    oSheet.Cells[38, 4] = "Immigration Status E.g. F1, F2 etc";
                    oSheet.Cells[38, 5] = "Primary Activity E.g. High School, Exchange etc";

                    //Row 39, 40, 41, 42, 43 is empty

                    //Attestation Statement Row 44
                    oSheet.Cells[44, 1] = "I attest the information provided here is accurate";
                    oSheet.Range[oSheet.Cells[44, 1], oSheet.Cells[44, 5]].Merge();
                    //Signature and Date Row 45
                    oSheet.Cells[45, 1] = "Signature";
                    oSheet.Cells[45, 4] = "Date";

                    //Row 46, 47, 48, 49, 50 will be empty
                    //For ISS use only heading Row 51
                    oSheet.Cells[51, 1] = "For ISS use only";

                    //Row 52 will be empty

                    //Reviewed By and Date - Row 53
                    oSheet.Cells[53, 1] = "Reviewed By";
                    oSheet.Cells[53, 4] = "Date";
                    #endregion

                    #region Bold Text for Labels
                    //Make the labels bold
                    for(int i=1;i<=53;i++)
                    {
                        oSheet.Cells[i, 1].Font.Bold = true;
                        oSheet.Cells[i, 4].Font.Bold = true;
                    }
                    //left out cells that are not part of for loop
                    oSheet.Cells[38, 2].Font.Bold = true;
                    oSheet.Cells[38, 5].Font.Bold = true;

                    
                    //underline the cells
                    oSheet.Cells[1, 1].Font.Underline = true;
                    oSheet.Cells[22, 1].Font.Underline = true;
                    #endregion
                    #region Underline and Borders For Values

                    for (int i = 2; i <= 45; i++)
                    {
                        //draw underlines for Street Address Fields
                        if (i == 10 || i == 11 || i == 16 || i == 17 || i == 26 || i == 27)
                        {
                            oSheet.Cells[i, 2].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            oSheet.Cells[i, 3].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            oSheet.Cells[i, 4].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            oSheet.Cells[i, 5].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        }
                        else if (i == 34 || i == 20)
                        {
                            oSheet.Cells[i, 5].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            continue;
                        }
                        //Blank Fields - Don't draw underlines
                        else if (i == 8 || i == 9 || i == 10 || i == 14 || i == 15 || i == 20 || i == 21 || i == 22 || i == 23 || i == 24 || i == 31 || i == 32 || i == 33 || i == 34 || i == 35 || i == 36 || i == 37 || i == 38 || i == 39 || i == 40 || i == 41 || i == 42 || i == 43 || i == 44)
                        {
                            continue;
                        }
                        //Label Value Fields need single cell underlines
                        oSheet.Cells[i, 2].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        oSheet.Cells[i, 5].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;



                    }
                    //Border for Immigration History
                    oRng = oSheet.get_Range("A36", "E43");
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng = oSheet.get_Range("A36", "E43");
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    //Outside only Borders for ISS Signature
                    oRng = oSheet.get_Range("A51", "E53");
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
                    oRng.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
                    #endregion

                    #region Values From Excel Data
                    ArrayList al = ExcelData.ExcelDataValuesCumulative.ElementAt(j);
                    oSheet.Cells[2, 2] = al[0];
                    oSheet.Cells[2, 5] = al[1];
                    oSheet.Cells[3, 2] = al[2];
                    oSheet.Cells[3, 5] = al[3];
                    oSheet.Cells[4, 2] = al[4];
                    oSheet.Cells[4, 5] = al[5];
                    oRng = oSheet.get_Range("B5");
                    oRng.NumberFormat = "DD-MMM-YYYY";
                    oSheet.Cells[5, 2] = al[6];
                    oSheet.Cells[5, 5] = al[7];
                    oSheet.Cells[6, 2] = al[8];
                    oSheet.Cells[6, 5] = al[9];
                    oSheet.Cells[7, 2] = al[10];
                    oSheet.Cells[7, 5] = al[11];
                    oSheet.Cells[10, 2] = al[12];
                    oSheet.Cells[11,2] = al[13];
                    oSheet.Cells[12, 2] = al[14];
                    oSheet.Cells[12, 5] = al[15];
                    oSheet.Cells[13, 2] = al[16];
                    oSheet.Cells[13, 5] = al[17];
                    var admitTerm = al[18];
                    //Extract year and semester from it
                    var year= admitTerm.ToString().Substring(0, 4);
                    var term = admitTerm.ToString().Substring(4);
                    string termText="";
                    if(term=="20")
                    {
                        termText = "Spring";
                    }
                    else if (term == "40")
                    {
                        termText = "Summer";
                    }
                    else if (term == "60")
                    {
                        termText = "Fall";
                    }
                    oSheet.Cells[1, 1] = "General Information Sheet (" + termText + " " + year + ")";
                    //Adding Barcode Image to sheet
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[55, 2];
                    b.IncludeLabel = true;
                    System.Drawing.Image img = b.Encode(BarcodeLib.TYPE.CODE39, al[3].ToString());
                    img.Save("C:\\images\\" + al[3] + ".jpg");
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float imageWidth = 327;
                    const float imageHeight = 139;
                    oSheet.Shapes.AddPicture("C:\\images\\" + al[3] + ".jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, imageWidth, imageHeight);
                    #endregion
                    string fileName = al[0]+", "+al[1];
                   // oSheet.Cells[56, 4] = al[3];
                   // oRange = oSheet.Cells[56, 4];
                    //oRange.Cells.Font.Size = 7;
                    oSheet.PageSetup.Zoom=92;
                    oSheet.PageSetup.FitToPagesWide = 1;
                    oWB.SaveAs(textBox1.Text+ "\\"+fileName+".xlsx");
                    
                    oWB.Close();
                    oXL.Quit();
                }
                catch (Exception ex)
                {
                    
                }
            }
            MessageBox.Show("Process Completed Successfully");
        }

        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbDataLoad.Value = e.ProgressPercentage;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                var worker = sender as BackgroundWorker;
               // worker.ReportProgress();
                ExcelData exceldata = new ExcelData();
                this.dvPreviewData.DataContext = exceldata;
            });
           
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
           
            MessageBox.Show("Data Load Complete");
        }
    }
}
