using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace apiDbExcel
{
    class exportToExcel
    {
        public static string connectString = @"Data Source=SEAN\MSSQL_SEAN;Initial Catalog = mydb; Integrated Security = True";
        //public string connectString = @"Data Source = NUMERAXIAL; Initial Catalog = Numerxial_Calculation; user=sa;Password=mnipl-1234";
        public exportToExcel(string ticker)
        {
            GenerateExcel(ticker);
        }

        public void GenerateExcel(string ticker)
        {
            System.Data.DataTable dtTable = new System.Data.DataTable("mydb.dbo.stock_historical");
            var list = new List<Dictionary<string, object>>();
            using (SqlConnection con = new SqlConnection(connectString))
            {
                String query_Insert = "INSERT into mydb.dbo.stock_table(stock_name, openP,highP,lowP)values ('BBBB', 49.74, 50.8, 48.99)";
                String query_Select = "select stock_name, openP, highP, lowP from stock_table";
                String query_SelectHis = "select ticker, stock_date, open_market, high, low, close_market, volumn FROM stock_historical ";

                String query_SelectAAA = "SELECT ticker, stock_date, open_market, high, low, close_market, volumn FROM stock_historical where ticker = 'AAPL'";

                String query_SelectBBB = "SELECT ticker, stock_date, close_market FROM stock_historical where ticker = 'A' ";
                string query_myQuery = "SELECT * from dbo.tlb_BABALog";
                using (SqlCommand cmd = new SqlCommand(query_myQuery, con))
                {
                    con.Open();
                    SqlDataAdapter adp = new SqlDataAdapter(cmd);
                    adp.Fill(dtTable);

                    JavaScriptSerializer serializer = new JavaScriptSerializer();



                    foreach (DataRow row in dtTable.Rows)
                    {
                        var dict = new Dictionary<string, object>();
                        foreach (DataColumn col in dtTable.Columns)
                        {
                            dict[col.ColumnName] = (Convert.ToString(row[col]));
                        }
                        list.Add(dict);
                    }
                    String outputString = serializer.Serialize(list);
                    Console.WriteLine(serializer.Serialize(list));
                    //Console.WriteLine(outputString);

                    //writeToFile(outputString);
                    //Console.WriteLine(list.Count);

                    // Console.WriteLine(String.Format("Stock Name" + "\t" + "Date" + "\t" + "Open Price" + "\t" + "High Price" + "\t" + "Low Prince" + "\t" + "Close Price" + "\t" + "Volumn" + "\t"));
                    /* foreach (DataRow row in dtTable.Rows)
                     {
                         //Console.WriteLine(String.Format(row[0] + "\t" + row[1] + "\t\t" + row[2] + "\t\t" + row[3] + "\t" +row[4] +"\t" + row[5] + "\t"+row[6]));
                     }*/
                    con.Close();
                    //writeToFile(outputString);
                }

                excelFile(dtTable, list);


            }
        }
            private static void excelFile(System.Data.DataTable dtTable, List<Dictionary<string, object>> list)
            {
                // File save dialog
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Execl files (*.xls)|*.xls";

                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.FileName = null;
                saveFileDialog.Title = "Save path of the file to be exported";


                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                /*
                string filepath = AppDomain.CurrentDomain.BaseDirectory;
                Console.WriteLine(filepath);
                string filename = @"output1.xlsx";
                Console.WriteLine(filepath + filename);*/


                // Microsoft.Office.Interop.Excel.Workbook wkbook = null;


                var wkbooks = xlApp.Workbooks;
                var wkbook = wkbooks.Add(Missing.Value);
                var wksheet = wkbook.ActiveSheet;

                Console.WriteLine(list[0].Keys);
                wksheet.Name = "Sean";

                try
                {
                    for (var i = 0; i < dtTable.Columns.Count; i++)
                    {
                        wksheet.Cells[1, i + 1] = dtTable.Columns[i].ColumnName;
                    }

                    //rows
                    for (var i = 0; i < dtTable.Rows.Count; i++)
                    {
                        for (var j = 0; j < dtTable.Columns.Count; j++)
                        {
                            wksheet.Cells[i + 2, j + 1] = dtTable.Rows[i][j];
                        }
                    }

                    //System.IO.FileInfo fileInfo = new System.IO.FileInfo(filename);

                    //File.SetAttributes(filename, ~FileAttributes.ReadOnly);
                    //File.SetAttributes(filename, ~FileAttributes.Hidden);


                    //wkbooks.Close();
                    wkbook.Close(true, Missing.Value, Missing.Value);
                    xlApp.Quit();

                    //wkbook.SaveAs(saveFileDialog, XlFileFormat.xlExcel8,false,
                    //false, false,false, XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    //Type.Missing, Type.Missing,Type.Missing, Type.Missing);

                    /*
                    wkbooks = null;
                    wkbook = null;
                    wksheet = null;
                    xlApp = null;*/

                    Marshal.ReleaseComObject(xlApp);
                    Marshal.ReleaseComObject(wkbooks);
                    Marshal.ReleaseComObject(wkbook);
                    Marshal.ReleaseComObject(wksheet);


                    //GC.Collect();

                }
                catch (Exception e)
                {

                    Console.WriteLine(e.ToString());
                }


            
        }
    }
}
