using System;
using System.Data.SqlClient;


namespace SeeDatabase
{
    class AdminTable
    {
        string connectString = @"Data Source=SEAN\MSSQL_SEAN;Initial Catalog = mydb; Integrated Security = True";
        //string connectString = @"Data Source = NUMERAXIAL; Initial Catalog = Numerxial_Calculation; user=sa;Password= mnipl-1234";
        string[] myQuery = new string[3];

        private string ticker = "";

        //constructor
        public AdminTable()
        {

        }

        //getter
        public string getTicker(string ticker)
        {
            return ticker;
        }

        //setter
        public void setTicker(string ticker)
        {
            this.ticker = ticker;
        }



        public void Create(string ticker)
        {
            using (SqlConnection con = new SqlConnection(connectString))
            {

                //string query_CreateMaster = @"CREATE TABLE tlb_master(ticker varchar(5) NOT NULL, stock_name varchar(63), PRIMARY KEY(ticker), UNIQUE(ticker))";
               // myQuery[0] = @"INSERT INTO  tlb_mastertb (ticker) values ('" + ticker + @"')";
                myQuery[0] = @"CREATE TABLE tlb_" + ticker + @"Log (" + ticker + @"LogID INT identity primary key,stock_value float NOT NULL,RecordOn date unique,RecordBy varchar(16))";
                myQuery[1] = @"CREATE TABLE tlb_" + ticker + @"RiskCalc(" + ticker + @"RiskCalcID INT identity primary key ,NormalVar float ,NPVar float,MarketRisk float,RecordOn date unique,RecordBy varchar(16), FOREIGN KEY(" + ticker + @"RiskCalcID) REFERENCES tlb_"+ticker+ @"Log(" + ticker + @"LogID))";
                myQuery[2] = @"CREATE TABLE tlb_" + ticker + @"LogCalc(" + ticker + @"LogCalcID INT identity primary key , LogValue float,PercentileValue decimal(5,4), stdDeviation float, MeanValue float,RecordOn date unique,RecordBy varchar(16), Last" + ticker + @"LogID int, FOREIGN KEY(" + ticker + @"LogCalcID) REFERENCES tlb_" + ticker + @"Log(" + ticker + @"LogID))";


                for (int i = 0; i < myQuery.Length; i++)
                {
                    try
                    {
                       // SqlConnection conn = new SqlConnection(connectString);
                        con.Open();
                        SqlCommand cmd = new SqlCommand(myQuery[i], con);

                        cmd.ExecuteNonQuery();
                        Console.WriteLine("Table Created Successfully...");
                        con.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("exception occured while creating table:" + e.Message + "\t" + e.GetType());
                    }

                }
              con.Close();
            }
            
        }

        public void Drop(string ticker)
        {
            using (SqlConnection con = new SqlConnection(connectString))
            {


                myQuery[0] = "DROP TABLE " + "tlb_" + ticker + "RiskCalc";
                myQuery[1] = "DROP TABLE " + "tlb_" + ticker + "LogCalc"; 
                myQuery[2] = "DROP TABLE " + "tlb_" + ticker + "Log";
                //myQuery[3] = @"DELETE FROM tlb_mastertb WHERE ticker = '" + ticker + @"' ";
                for (int i = 0; i < myQuery.Length; i++)
                {
                    try
                    {
                        //SqlConnection conn = new SqlConnection(connectString);
                        SqlCommand cmd = new SqlCommand(myQuery[i], con);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        Console.WriteLine("Table Delete Successfully...");
                        con.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("exception occured while creating table:" + e.Message + "\t" + e.GetType());
                    }

                }
                con.Close();
            }
        }


    }
}
