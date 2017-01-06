using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WindowsFormsApplication1
{


    public partial class Form1 : Form
    {

        String excelFileName = "";   // Excel file
        Excel.Worksheet m_bookingsWorksheet;
        Excel.Application m_application;

        Excel._Workbook m_workbook;
        int m_bookings_totRows;

        Excel._Worksheet m_sheet1;
        int m_sheet1_totRows;

        DateTime m_usageMonthYear = DateTime.Now;

        public Form1()
        {
            InitializeComponent();
        }

       

        private void button1_Click(object sender, System.EventArgs e)
        {

           
          
            Excel.Range oRng;
            Excel.Range dateRange;



            try
            {
                if(excelFileName.Length == 0)
                {
                    MessageBox.Show("Please select an input file using the browse button.");
                    return;
                }

            

                //Start Excel and get Application object.
                m_application = new Excel.Application();
                m_application.Visible = true;

                //Get a new workbook.
                //         oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));

                m_workbook = m_application.Workbooks.Open(excelFileName);

                //       oSheet = (Excel._Worksheet)m_workbook.ActiveSheet;  // need Sheet1, bookings might be first and active if saved sheet
                m_sheet1 = (Excel.Worksheet)m_workbook.Worksheets["Sheet1"];
                //Add table headers going cell by cell.
                //      oSheet.Cells[1, 1] = "First Name";
                //     oSheet.Cells[1, 2] = "Last Name";
                //     oSheet.Cells[1, 3] = "Full Name";
                //     oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                //     oSheet.get_Range("A1", "D1").Font.Bold = true;
                //    oSheet.get_Range("A1", "D1").VerticalAlignment =
                //       Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                //      oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //      oRng = oSheet.get_Range("C2", "C6");
                //     oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                //     oRng = oSheet.get_Range("D2", "D6");
                //     oRng.Formula = "=RAND()*100000";
                //     oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                //     oRng = oSheet.get_Range("A1", "D1");
                //     oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                //     DisplayQuarterlySales(oSheet);


                // get Usage Period 

                //   dateRange = oSheet.get_Range("AD2","AD2").Value2;
                string sDate = m_sheet1.Cells[2, 30].Value2.ToString(); 
     //           string sDate = (dateRange.Cells[1, 1]).Value2.ToString();

                double date = double.Parse(sDate);

                //  var dateTime = DateTime.FromOADate(date).ToString("MMMM dd, yyyy");

                m_usageMonthYear = DateTime.FromOADate(date);

                createBookingsSheet(m_workbook);

                // Now that booking sheet is created.
                // Add provider_amount composite helper column to Bookings Sheet

                Excel.Range rng = m_bookingsWorksheet.get_Range("A1", Missing.Value);

                rng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                                        Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                m_bookingsWorksheet.Cells[1, 1] = "providerid_grossamount";

                // Add relative formula
                // Get the used Range
        //        Excel.Range usedRange = m_bookingsWorksheet.UsedRange;
                m_bookings_totRows = m_bookingsWorksheet.UsedRange.Rows.Count;
                oRng = m_bookingsWorksheet.get_Range("A2", "A" + m_bookings_totRows);
                oRng.Formula = "=B2 & D2";



                // Add Rebate column to Sheet1

                Excel.Range orng = m_sheet1.get_Range("AC1", Missing.Value);
                orng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                              Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                oRng.NumberFormat = "$0.00";
                m_sheet1.Cells[1, "AC"] = "Rebate";
   

                // Add vlookup function
                // Get the used Range
                //      Excel.Range oSheetUsedRange = m_sheet1.UsedRange;
                m_sheet1_totRows = m_sheet1.UsedRange.Rows.Count;
                oRng = m_sheet1.get_Range("AC2", "AC" + m_sheet1_totRows);
                oRng.Formula = "=IFERROR(VLOOKUP(D2&AB2,Bookings!$A$2:$E$" + m_bookings_totRows +",5,FALSE),0)";  // $ means absolute cell not relative in formula



                // Add Totals to Both Sheets



                //These two lines do the magic.
                // m_bookingsWorksheet.Columns.ClearFormats();
                //   m_bookingsWorksheet.Rows.ClearFormats();

                // int iTotalColumns = m_bookingsWorksheet.UsedRange.Columns.Count;
                //   int iTotalRows = m_bookingsWorksheet.UsedRange.Rows.Count;

                Excel.Range dcolrng = m_bookingsWorksheet.get_Range("D2", "D" + m_bookings_totRows);
                dcolrng.NumberFormat = "###,###,##0.00";

                String Expr1 = "=SUM(D2:D" + m_bookings_totRows + ")";
                m_bookingsWorksheet.Cells[(m_bookings_totRows + 4), "D"] = Expr1;
                //          m_bookingsWorksheet.get_Range("D" + (iTotalRows + 4), "D" + (iTotalRows + 4)).Cells[1, 1] = Expr3;
                //     m_bookingsWorksheet.get_Range("C" + (iTotalRows + 4), "D" + (iTotalRows + 4)).Cells[1, 1] = "Gross Total";


                Excel.Range ecolrng = m_bookingsWorksheet.get_Range("E2", "E" + m_bookings_totRows);
                ecolrng.NumberFormat = "###,###,##0.00";

                // Add  Gross Bookings total
                //       String Expr2 = "=SUMIF(C2:C" + iTotalRows + ", \"Gross Bookings\"  ,D2:D" + iTotalRows + ")";
                String Expr2 = "=SUM(E2:E" + m_bookings_totRows + ")";
                m_bookingsWorksheet.Cells[(m_bookings_totRows + 4), "E"] = Expr2;
                //    m_bookingsWorksheet.get_Range("D" + (iTotalRows + 3), "D" + (iTotalRows + 3)).Cells[1, 1] = Expr2;
                //   m_bookingsWorksheet.get_Range("C" + (iTotalRows + 3), "D" + (iTotalRows + 3)).Cells[1, 1] = " Rebate Total";


                Excel.Range aacolrng = m_sheet1.get_Range("AA2", "AA" + m_sheet1_totRows);
                aacolrng.NumberFormat = "###,###,##0.00";

                String Expr3 = "=SUM(AA2:AA" + m_sheet1_totRows + ")";
                m_sheet1.Cells[(m_sheet1_totRows + 4), "AA"] = Expr3;

                Excel.Range accolrng = m_sheet1.get_Range("AC2", "AC" + m_sheet1_totRows);
                accolrng.NumberFormat = "###,###,##0.00";

                String Expr4 = "=SUM(AC2:AC" + m_sheet1_totRows + ")";
                m_sheet1.Cells[(m_sheet1_totRows + 4), "AC"] = Expr4;



                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                m_application.Visible = true;
                m_application.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void createBookingsSheet(Excel._Workbook m_objBook)
        {

            // Start a new workbook in Excel.
            // m_objExcel = new Excel.Application();
            // m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            // m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            // check for existing bookings sheet
            Boolean found = false;
            foreach (Excel.Worksheet displayWorksheet in m_objBook.Worksheets)
            {
                if (displayWorksheet.Name.Equals("Bookings"))
                {
                    found = true;
                    m_bookingsWorksheet = displayWorksheet;
                }
            }
            if (!found)
            {
                // Add new Sheet for Bookings
                m_bookingsWorksheet = (Excel.Worksheet)m_workbook.Worksheets.Add();
                m_bookingsWorksheet.Name = "Bookings";

                // Create a QueryTable that starts at cell A1.
                //    Excel.Sheets objSheets = (Excel.Sheets)m_objBook.Worksheets;
                //    Excel._Worksheet objSheet = (Excel._Worksheet)(objSheets.get_Item(1));

                Excel.Range objRange = m_bookingsWorksheet.get_Range("A1");
                Excel.QueryTables objQryTables = m_bookingsWorksheet.QueryTables;

                // Form query for billing database adding 1 month to retrieve corresponding billing data

                    DateTime queryDate = (m_usageMonthYear.AddMonths(1));
             //   DateTime queryDate = m_usageMonthYear;
                var startOfMonth = new DateTime(queryDate.Year, queryDate.Month, 1);

                //    string SQLStr = "SELECT partnerid, orderidpm, enteredby, sum(detailTotalLC) " +
                //                    "FROM Bookings.dbo.[odbc-acctg-vspp-detailed]" +
                //         "where Month = '20161101' " +
                //                   "where Month = '" + startOfMonth.ToString("yyyyMMdd") +
                //                  "' group by partnerid, orderidpm, enteredby";

                string SQLStr = "SELECT partnerid, orderidpm," +
                              " SUM( CASE enteredby " +
                                   " WHEN 'Gross Bookings' " +
                                   " THEN detailTotalLC" +
                                   " ELSE 0 " +
                              " END) as totalGross, " +
                              " SUM( CASE " +
                                   " WHEN enteredby !=  'Gross Bookings' " +
                                   " THEN detailTotalLC " +
                                   " ELSE 0 " +
                                   " END) as totalRebate, " +
                                   " SUM(qty) as points " +
                              " FROM Bookings.dbo.[odbc-acctg-vspp-detailed]" +
                              " where Month = '" + startOfMonth.ToString("yyyyMMdd") +
                              "' group by partnerid, orderidpm";


                Excel._QueryTable objQryTable = (Excel._QueryTable)objQryTables.Add("ODBC;DSN=store-dbrepo1", objRange, SQLStr);
                objQryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;
                objQryTable.Refresh(false);  // do not run in background




            }
            else {
                String sMsg = "Billing Data exists, refresh data from database?";
                DialogResult iRet = MessageBox.Show(sMsg, "Question", MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2,(MessageBoxOptions)0x40000);
                if (iRet == DialogResult.Yes)
                {

                    // get query table
                    Excel.Range objRange = m_bookingsWorksheet.get_Range("A1");
                    Excel.QueryTables objQryTables = m_bookingsWorksheet.QueryTables;
                    foreach (Excel.QueryTable qt in objQryTables) {
                        qt.Refresh();
                    }
                }


                // Add a Gross Bookings Sum and Rebate Sum

                //Get the used Range
                //   Excel.Range usedRange = m_bookingsWorksheet.UsedRange;
                //       int totRows = m_bookingsWorksheet.UsedRange.Rows.Count;

                //Iterate the rows in the used range
                //       double grossSum = 0;
                //       double rebateSum = 0;

                //      for(int i=2; i< totRows-2;i++) 
                //      {

                //          String v = (String)  m_bookingsWorksheet.get_Range("C" + i, "C" + i).Cells[1, 1].Value2;
                //Do something with the row.
                // Values are stored in rows not columns
                //         if (String.Equals(v, @"Gross Bookings", StringComparison.OrdinalIgnoreCase))
                //        {
                // grossSum =  grossSum + int.Parse(row.Cells[1, 4].ToString());
                //      grossSum = int.Parse((string)(m_bookingsWorksheet.Cells[10, 2] as Excel.Range).Value);
                //           var v2 = m_bookingsWorksheet.get_Range("D" + i, "D" + i).Cells[1, 1].Value2;
                //     if (v2.GetType() != typeof(Double)) ;
                //           if (v2 != null)
                //          {
                //             String myVal = (string)v2.ToString();
                //            double rebate;
                //           if (Double.TryParse(myVal, out rebate))
                //              grossSum = grossSum + rebate;
                //         else
                //            Console.WriteLine("{0} is outside the range of a Double.",
         //       v2);
                //    }
                //    }
                //   else {
                //      String s2 = (String) (m_bookingsWorksheet.Cells[4,i] as Excel.Range).Value2;
                //      String s2 = (string)(excelWorksheet.Cells[10, 2] as Excel.Range).Value;
                // cannot assume value is string maybe null                  var v3 = m_bookingsWorksheet.get_Range("D" + i, "D" + i).Cells[1, 1].Value2.ToString();

                //      ((Excel.Range)xlSheet.Cells[i, "B"]).Value2 != null) 

                //     if (m_bookingsWorksheet.get_Range("D" + i, "D" + i).Cells[1, 1].Value2 != null) { }
                //     var v3 = m_bookingsWorksheet.get_Range("D" + i, "D" + i).Cells[1, 1].Value2;
                
                //    if (v3 != null) {
                //          rebateSum = rebateSum + double.Parse(v3);
                //        String myVal = (string)v3.ToString();
                //       double rebate;
                //   if (Double.TryParse(myVal, out rebate))
                //       rebateSum = rebateSum + rebate;
                //   else
                //       Console.WriteLine("{0} is outside the range of a Double.",
         //       v3);
                //   }
                //       Console.Write(".");
                //   }

     //       }

            // How to read a date value
            //   double d = double.Parse(b);
            //   DateTime conv = DateTime.FromOADate(d);
            // or
            //   string sDate = (xlRange.Cells[4, 3] as Excel.Range).Value2.ToString();

            //    double date = double.Parse(sDate);

            //   var dateTime = DateTime.FromOADate(date).ToString("MMMM dd, yyyy");



            //       m_bookingsWorksheet.Cells[3, totRows + 1].Value2 = "Total Gross";
            //       m_bookingsWorksheet.Cells[4, totRows + 1].Value2 = grossSum;

            //       m_bookingsWorksheet.Cells[3, totRows + 2].Value2  = "Total Rebate";
            //       m_bookingsWorksheet.Cells[4, totRows + 2].Value2 = rebateSum;


         

            }



            // Save the workbook and quit Excel.
            //     objBook.SaveAs(m_strSampleFolder + "Book4.xls", objOpt, objOpt,
            //         objOpt, objOpt, objOpt, Excel.XlSaveAsAccessMode.xlNoChange, m_objOpt, m_objOpt,
            //          objOpt, objOpt);
            //    m_objBook.Close(false, objOpt, m_objOpt);
            //      m_objExcel.Quit();
        }

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;

            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                    MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
                = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
                = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
                Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
                Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
       
            int size = -1;
          
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {
                excelFileName = openFileDialog1.FileName;
                this.label2.Text = excelFileName;

                //  try
                //     {
                //         string text = File.ReadAllText(file);
                //         size = text.Length;
                //     }
                //     catch (IOException)
                //     {
                //      }
                //   }
                //       Console.WriteLine(size); // <-- Shows file size in debugging mode.
                //       Console.WriteLine(result); // <-- For debugging use.
            }
        
        }

   //     private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
   //     {
   //         m_usageMonthYear = dateTimePicker1.Value;
   //    }
    }
}
