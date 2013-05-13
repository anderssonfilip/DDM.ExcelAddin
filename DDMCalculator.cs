using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace DDM.ExcelAddIn
{
    public class DDMCalculator
    {
        private readonly decimal _requiredRateOfReturn;

        public DDMCalculator(decimal requiredRateOfReturn)
        {
            _requiredRateOfReturn = requiredRateOfReturn;
        }

        public DDM Calculate(List<Dividend> dividends)
        {
            var ddm = new DDM();
            var annualDividends = new List<decimal>();

            annualDividends.Add(0m);

            var dividendFrequency = 0; 
            var currentYearDividendPayments = 0;

            if (dividends.First().Date > dividends.Last().Date)
            {
                dividends.Reverse();
            }

            var year = dividends.First().Date.Year;

            foreach(var dividend in dividends)
            {
                if (dividend.Date.Year == year)
                {
                    dividendFrequency += 1;
                }
                else if(dividend.Date.Year > year)
                {
                    year = dividend.Date.Year;
                    annualDividends.Add(0m);

                    if (year == DateTime.Now.Year)
                    {
                        currentYearDividendPayments += 1;
                    }
                    else
                    {
                        dividendFrequency = 1;
                    }
                }

                annualDividends[annualDividends.Count-1] += dividend.Amount;
            }

            if (currentYearDividendPayments != 0m && dividendFrequency != currentYearDividendPayments)
            {
                annualDividends[annualDividends.Count - 1] = annualDividends.Last() * (dividendFrequency / currentYearDividendPayments);
                ddm.IsAnnualDividendAnnualized = true;
            }

            ddm.D = annualDividends.Last();

            var growthRates = new List<decimal>(annualDividends.Count);
            for (var i = 1; i < annualDividends.Count; i++)
            {
                growthRates.Add((annualDividends[i] / annualDividends[i - 1]) - 1);
            }

            ddm.G = growthRates.Count > 0 ? growthRates.Average() : 0;
            ddm.R = _requiredRateOfReturn;
            return ddm;
        }
    }
}
