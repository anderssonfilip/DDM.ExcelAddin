using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace DDM.ExcelAddIn
{
    public partial class InvestRibbon
    {
        private void InvestRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var win = new DDMWindow();

            win.Show();

     
        }      
    }
}
