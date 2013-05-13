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

                WriteDDMValues(ddm, dividends);
            }
        }

        internal void WriteDDMValues(DDM ddm, List<Dividend> dividends)
        {
            var xlApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            var xlWorkbook = xlApp.ActiveWorkbook;
            var xlWorksheet = xlWorkbook.ActiveSheet;

            Range selectedRange = xlApp.ActiveWindow.RangeSelection;

            var extraRows = writeDividends.IsChecked == true ? dividends.Count : 0;

            if (selectedRange.Rows.Count != 5 + extraRows || selectedRange.Columns.Count != 2)
            {
                MessageBoxResult mbResult;

                do
                {
                    mbResult = MessageBox.Show(string.Format("Select a range with size {0}:{1} to write result", 5 + extraRows, 2), "Select range", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    selectedRange = xlApp.ActiveWindow.RangeSelection;

                    if (mbResult == MessageBoxResult.OK && selectedRange.Rows.Count == (5 + extraRows) && selectedRange.Columns.Count == 2)
                        break;
                }
                while (mbResult != MessageBoxResult.Cancel);

                if (mbResult == MessageBoxResult.Cancel)
                    return;
            }

            if (selectedRange.Rows.Count == 5 + extraRows && selectedRange.Columns.Count == 2)
            {
                ((Range)selectedRange.Cells[1, 1]).Value = "Ticker";
                ((Range)selectedRange.Cells[1, 2]).Value = ticker;

                ((Range)selectedRange.Cells[2, 1]).Value = "P";
                selectedRange.Cells[2, 2] = string.Format("={0}/({1}-{2})", selectedRange.Cells[3, 2].Address, selectedRange.Cells[4, 2].Address, selectedRange.Cells[5, 2].Address);

                selectedRange.Cells[3, 1] = "D";
                selectedRange.Cells[3, 2] = ddm.D;

                selectedRange.Cells[4, 1] = "r";
                selectedRange.Cells[4, 2].Value2 = ddm.R;
                selectedRange.Cells[4, 2].NumberFormat = "#.##%";

                selectedRange.Cells[5, 1].Value = "g";
                selectedRange.Cells[5, 2].Value = ddm.G;
                selectedRange.Cells[5, 2].NumberFormat = "#.##%";

                if (writeDividends.IsChecked == true)
                {
                    int row = 6;

                    foreach (var dividend in dividends)
                    {
                        selectedRange.Cells[row, 1].Value = dividend.Date;
                        selectedRange.Cells[row, 2].Value = dividend.Amount;
                        row++;
                    }
                }
            }
        }
    }
}
