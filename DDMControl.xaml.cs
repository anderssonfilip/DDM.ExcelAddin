using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;

namespace DDM.ExcelAddIn
{
    public partial class DDMControl : UserControl
    {
        private string ticker;
        private int year;
        private decimal rrr;

        public DDMControl()
        {
           InitializeComponent();

           yearBox.Text = DateTime.Now.AddYears(-10).Year.ToString();
           rBox.Text = 0.15.ToString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ticker = tickerBox.Text;

            if (!string.IsNullOrEmpty(ticker) && int.TryParse(yearBox.Text, out year) && decimal.TryParse(rBox.Text, out rrr))
            {
                var dividendLoader = new DividendLoader();

                var dividends = dividendLoader.GetDividendHistory(tickerBox.Text, year);

                var ddmCalculator = new DDMCalculator(rrr);

                var ddm = ddmCalculator.Calculate(dividends);

                WriteDDMValues(ddm);

                if (writeDividends.IsChecked == true)
                {
                    WriteDividendPayouts(dividends);
                }
            }
        }


        private void WriteDDMValues(DDM ddm)
        {
            var xlApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            var xlWorkbook = xlApp.ActiveWorkbook;
            var xlWorksheet = xlWorkbook.ActiveSheet;

            ((Range)xlWorksheet.Cells[1, 1]).Value2 = "Ticker";
            ((Range)xlWorksheet.Cells[1, 2]).Value = ticker;

            ((Range)xlWorksheet.Cells[2, 1]).Value = "P";
            xlWorksheet.Cells[2, 2] = ddm.P;

            xlWorksheet.Cells[2, 1] = "D";
            xlWorksheet.Cells[2, 2] = ddm.D;

            xlWorksheet.Cells[3, 1] = "R";
            xlWorksheet.Cells[3, 2].Value2 = ddm.R;

            xlWorksheet.Cells[4, 1].Value = "G";
            xlWorksheet.Cells[4, 2].Value = ddm.G;


        }

        private void WriteDividendPayouts(List<Dividend> dividends)
        {
            foreach (var dividend in dividends)
            {

            }
        }
    }
}
