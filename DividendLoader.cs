using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace DDM.ExcelAddIn
{
    public class DividendLoader
    {
        // g=v sets the flag get dividend payout and date
        private const string Url = "http://ichart.finance.yahoo.com/table.csv?s={0}&a={1}&b={2}&c={3}&d={4}&e={5}&f={6}&g=v&ignore.csv";

        public static string GetResponse(string url)
        {
            int num = 0;
            while (num <= 3)
            {
                using (var client = new WebClient())
                {
                    try
                    {
                        return client.DownloadString(url);
                    }
                    catch (Exception exception)
                    {
                        //Use your logging utility of choice here....
                    }
                    continue;
                }
            }
            return "";
        }


        public List<Dividend> GetDividendHistory(string code, int fromYear)
        {
            var dividends = new List<Dividend>();

            var fromDate = new DateTime(fromYear, 1, 1);
            var now = DateTime.Now;
            if (fromDate >= now)
            {
                return dividends;
            }

            var url = string.Format(Url,
                                new object[]
                                        {
                                            code, fromDate.Month, fromDate.Day, fromDate.Year, now.Month,  now.Day, now.Year
                                        });

            const int index = 0;
            var provider = new CultureInfo("en-US", true);
            var response = GetResponse(url);
            try
            {

                using (var reader = new StringReader(response))
                {
                    reader.ReadLine();
                    while (reader.Peek() > -1)
                    {
                        var strArray = reader.ReadLine().Split(new char[] { ',' });
                        var date = DateTime.Parse(strArray[index].Replace("\"", ""), provider);
                        if (date < fromDate) continue;
                        var dividend = new Dividend
                        {
                            Date = date,
                            Amount = Convert.ToDecimal(strArray[1 + index])
                        };
                        dividends.Add(dividend);
                    }
                }
            }
            catch (Exception exception)
            {
                //Use your logging utility of choice here....
            }
            return dividends;
        }

    }
}
