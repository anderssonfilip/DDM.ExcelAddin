using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace DDM.ExcelAddIn.Test
{
    [TestClass]
    public class DDM_ExcelAddIn_Tests
    {
        [TestMethod]
        public void GetResponseTest()
        {
            var response = DividendLoader.GetResponse("http://www.blankwebsite.com/");

            Assert.IsFalse(string.IsNullOrEmpty(response));

        }

        [TestMethod]
        public void GetDividends_For_MCD_Since_2000_Test()
        {
            var dl = new DividendLoader();

            var dividends = dl.GetDividendHistory("MCD", 2000);

            Assert.IsTrue(dividends.Count > 0);

            var date = DateTime.Now.Date;

            foreach (var dividend in dividends)
            {
                Assert.IsTrue(dividend.Date < date);
                Assert.IsTrue(dividend.Amount >= 0);

                date = dividend.Date;
            }
        }

        [TestMethod]
        public void DividendCalculator_5Quarters_Test()
        {
            var dc = new DDMCalculator(0.12m);

            var now = new DateTime(DateTime.Now.Year, 2, 28);

            var result = dc.Calculate(new List<Dividend>{
                new Dividend{Date = now.AddMonths(-3*4), Amount = 0.45m},
                new Dividend{Date = now.AddMonths(-3*3), Amount = 0.45m},
                new Dividend{Date = now.AddMonths(-3*2), Amount = 0.45m},
                new Dividend{Date = now.AddMonths(-3), Amount = 0.45m},
                new Dividend{Date = now, Amount = 0.5m},
            });

            Assert.IsTrue(result.IsAnnualDividendAnnualized);
            Assert.AreEqual(2m, result.D);
            Assert.AreEqual(0.1111m, Math.Round(result.G, 4));
            Assert.AreEqual(0.12m, result.R);
            Assert.AreEqual(225m, Math.Round(result.P));
        }


        [TestMethod]
        public void DividendCalculator_IncorrectDateOrder_Test()
        {
            var r = 0.10m;

            var dc = new DDMCalculator(r);

            var now = new DateTime(DateTime.Now.Year, 2, 28);

            var result = dc.Calculate(new List<Dividend>{
                new Dividend{Date = now, Amount = 1m},
                new Dividend{Date = now.AddMonths(-3*4), Amount = 1m}
            });

            Assert.IsFalse(result.IsAnnualDividendAnnualized);
            Assert.AreEqual(2m, result.D);
            Assert.AreEqual(0m, result.G);
            Assert.AreEqual(r, result.R);
            Assert.AreEqual(20m, result.P);
        }
    }
}
