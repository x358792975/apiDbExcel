using apiDbExcel;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace ApiToDatabase
{
    class getApiCall
    {
        public getApiCall(string ticker) {
            ConnectToApi(ticker); 
        }

        public void ConnectToApi(string ticker) {
            string pre = @"https://api.intrinio.com/prices?identifier=";
            //string ticker = "C";
            string post = @"&start_date=1900-01-01&page_number=";
            int currentPage = 1;

            string url = pre + ticker + post + currentPage;

            string url2 = @"https://api.intrinio.com/prices?identifier=SNAP&start_date=1900-01-01&page_number=" + currentPage+ @"";


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);


            request.Method = "GET";
            request.Headers["Authorization"] = "Basic " + Convert.ToBase64String
                    (Encoding.Default.GetBytes("c248b72786804ad428d98b29bef7c1c3:a0db0996cdc891c8d7732a9fbf6803c1"));

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream receiveStream = response.GetResponseStream();
            string jsonString = "";

                // Pipes the stream to a higher level stream reader with the required encoding format. 
            StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
            jsonString += readStream.ReadToEnd();

            response.Close();
            readStream.Close();
            currentPage++;

            var obj = new JavaScriptSerializer().Deserialize<Rootobject>(jsonString);
            int totalPage = obj.total_pages;
            new connectToDb(ticker,obj);
            while(currentPage<= totalPage)
            {
  
                string urltemp = pre + ticker + post + currentPage;


                HttpWebRequest requesttemp = (HttpWebRequest)WebRequest.Create(urltemp);

                requesttemp.Method = "GET";
                requesttemp.Headers["Authorization"] = "Basic " + Convert.ToBase64String
                    (Encoding.Default.GetBytes("c248b72786804ad428d98b29bef7c1c3:a0db0996cdc891c8d7732a9fbf6803c1"));

                HttpWebResponse responsetemp = (HttpWebResponse)requesttemp.GetResponse();
                Stream receiveStreamtemp = responsetemp.GetResponseStream();

                // Pipes the stream to a higher level stream reader with the required encoding format. 
                StreamReader readStreamtemp = new StreamReader(receiveStreamtemp, Encoding.UTF8);
                
                string jsonStr2 = "";
                jsonStr2 = readStreamtemp.ReadToEnd();
                    //Console.WriteLine(jsonString);
                responsetemp.Close();
                readStreamtemp.Close();
                Console.WriteLine("Current Page is " + currentPage);
                var obj2 = new JavaScriptSerializer().Deserialize<Rootobject>(jsonStr2);
                new connectToDb(ticker, obj2);
                currentPage++;
                    //writeToFile(jsonString);
                
            }
            Console.WriteLine("Done");
           // new exportToExcel(ticker);
            Console.ReadLine();
        }
        private static void writeToFile(string outputString)
        {
            using (StreamWriter writetext = new StreamWriter("output.txt"))
            {
                writetext.WriteLine(outputString);
                Console.WriteLine("Wrote into the file");
            }

        }
    }
}

