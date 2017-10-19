using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace ApiToDatabase
{
    class connectToDb
    {
        //string connectString = @"Data Source = NUMERAXIAL; Initial Catalog = Numerxial_Calculation; user=sa;Password= mnipl-1234";
        string connectString = @"Data Source=SEAN\MSSQL_SEAN;Initial Catalog = mydb; Integrated Security = True";

        public connectToDb(string ticker,Rootobject obj)
        {

            DoConnect(ticker,obj);
        }

        public void DisplayTable()
        {
            using (SqlConnection con = new SqlConnection(connectString))
            {

                DataTable dtTable = new DataTable("mydb.dbo.tlb_mastertb");
            string myQuery = @"select * from dbo.tlb_mastertb";
            try
            {
                SqlConnection conn = new SqlConnection(connectString);
                SqlCommand cmd = new SqlCommand(myQuery, conn);
                conn.Open();
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                adp.Fill(dtTable);
                foreach (DataRow row in dtTable.Rows)
                {
                    Console.WriteLine(String.Format(row[0] + "\t" + row[1] + "\t\t" + row[2] + "\t\t" + row[3] + "\t\t"));
                }
                cmd.ExecuteNonQuery();
                // Console.WriteLine("Table Created Successfully...");
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("exception occured while creating table:" + e.Message + "\t" + e.GetType());
            }
        }

        }

        public void DoConnect(string ticker,Rootobject obj)
        {
            //string author = "Sean";
            foreach (var p in obj.data) {
                /*
                DateTime d = DateTime.Parse(p.date);
                String s = d.ToString(CultureInfo.CreateSpecificCulture("en-US"));*/

                string myQuery = @"insert into dbo.tlb_" + ticker + @"Log (stock_value, RecordOn, RecordBy) values ( "
                                    + p.close + @",  '" + Convert.ToDateTime(p.date) + @"', 'Sean')";

                string myQuery2 = @"Select * from dbo.tlb_" + ticker + @"Log ";
                
                //Console.WriteLine(Convert.ToDateTime(p.date));
                using (SqlConnection con = new SqlConnection(connectString))
                {
                    try
                    {
                        SqlConnection conn = new SqlConnection(connectString);
                        SqlCommand cmd = new SqlCommand(myQuery, conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        //Console.WriteLine("Data inserted...");
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("exception occured while creating table:" + e.Message + "\t" + e.GetType());
                    }

                 }
               
             }

            }
    }
}

