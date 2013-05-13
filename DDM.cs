using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DDM.ExcelAddIn
{
    public class DDM
    {
        // required rate of return
        public decimal R;

        // annual dividend growth reate
        public decimal G;

        // annual dividend
        public decimal D;

        public bool IsAnnualDividendAnnualized;

        // price according to DDM
        public decimal P
        {
            get
            {
                return R - G == decimal.Zero ? decimal.Zero : D / (R - G);
            }
        }
    }
}
