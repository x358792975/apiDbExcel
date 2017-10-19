using SeeDatabase;
using System;
using System.Data.SqlClient;

namespace readWriteFile
{


    class checkTables
    {

        //string connectString = @"Data Source = NUMERAXIAL; Initial Catalog = Numerxial_Calculation; user=sa;Password=mnipl-1234";
        string connectString = @"Data Source=SEAN\MSSQL_SEAN;Initial Catalog = mydb; Integrated Security = True";

        public checkTables()
        {

        }

        public void ForDelete(string ticker)
        {
            string tableName = "tlb_" + ticker + "Log";

            using (SqlConnection con = new SqlConnection(connectString))
            {
                string myQuery = @"select case when exists((select * from mydb.INFORMATION_SCHEMA.TABLES where TABLE_NAME = '" + tableName + @"' )) then 1 else 0 end";

                try
                {

                    SqlCommand command = new SqlCommand(myQuery, con);
                    con.Open();
                    int result = (Int32)command.ExecuteScalar();
                    //Console.WriteLine(result);
                    con.Close();

                    if (result == 1)
                    {
                        Console.WriteLine("Deleting tables.");
                        AdminTable adTb = new AdminTable();
                        adTb.Drop(ticker);
                        return;
                    }
                    // return true;
                    else
                    {
                        return;
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine("SQL Error " + ex.Message.ToString());
                    // return false;
                }
            }


        }

        public void ForAdd(string ticker)
        {
            string tableName = "tlb_"+ticker+"Log";

            using (SqlConnection con = new SqlConnection(connectString))
            {
            string myQuery = @"select case when exists((select * from mydb.INFORMATION_SCHEMA.TABLES where TABLE_NAME = '"+tableName+ @"' )) then 1 else 0 end";

                try
                {

                    SqlCommand command = new SqlCommand(myQuery, con);
                    con.Open();
                    int result = (Int32)command.ExecuteScalar();
                    //Console.WriteLine(result);
                    con.Close();

                    if (result == 1)
                    {
                        return;
                    }
                    // return true;
                    else
                    {
                        Console.WriteLine("Creating new tables for the new stock.");
                        AdminTable adTb = new AdminTable();
                        adTb.Create(ticker);
                    }
                    new ApiToDatabase.getApiCall(ticker);
                }
                catch (SqlException ex)
                {
                    Console.WriteLine("SQL Error " + ex.Message.ToString());
                   // return false;
                }
            }


            }

        }
    }

