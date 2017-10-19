using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace readWriteFile
{
    class readFile
    {
        public string addPath = @"C:/Users/SeanCui/Documents/Sample.txt";
        public string delPath = @"C:/Users/SeanCui/Documents/Delete.txt";
        public readFile()
        {
            ReadAdd();
            ReadDelete();
        }

        public void ReadDelete()
        {
            string line;
            try
            {
                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(delPath);

                //first line of text
                line = sr.ReadLine();



                //Continue to read until you reach end of file
                while (line != null)
                {

                    checkTables ckDel = new checkTables();
                    ckDel.ForDelete(line.ToUpper());
                    //Console.WriteLine(line);
                    //Read the next line
                    line = sr.ReadLine();
                }

                //close the file
                sr.Close();
                writeToFile(delPath);
                //Console.ReadLine();
                Console.WriteLine("Finished reading Delete File");

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }
        }

        public void ReadAdd()
        {
            string line;
            try
            {
                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(addPath);

                //first line of text
                line = sr.ReadLine();

                //Continue to read until you reach end of file
                while (line != null)
                {
                    checkTables ckAdd = new checkTables();
                    ckAdd.ForAdd(line.ToUpper());
                    //Read the next line
                    line = sr.ReadLine();
                }

                //close the file
                sr.Close();
                //writeToFile(addPath);
                Console.WriteLine("Finished reading Add File");
                //Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Done.");
            }
        }

        public void writeToFile(string fileName)
        {
            string line = "";
            File.WriteAllText(fileName, line);
        }
    }
}
