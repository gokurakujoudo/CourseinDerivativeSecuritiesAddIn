using System;
using System.Windows;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.Alpha
{
    public static class PlayGround
    {
        private static readonly Application App = (Application)ExcelDnaUtil.Application;

        [ExcelCommand(MenuText="set formula")]
        public static void SetFormula() {
            var formula = @"=Sum(A1:B5)";
            Range selection = App.Selection;
            if (selection != null) {
                selection.Formula = formula;
            }
        }

        [ExcelCommand(MenuText = "BisectionSearchTemp")]
        public static void BisectionSearchtTemp() {
            Range selection = App.Selection;
            if (selection == null) {
                MessageBox.Show("Invalid selection", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Range newr = selection.Resize[6, 2];
            newr.Value = new object[6, 2] {
                {"TYPE", "BISECTION_SEARCH"},
                {"X", 0},
                {"Y", "formula"},
                {"TARGET", 10},
                {"LOWER", 0},
                {"UPPER", 15}
            };
            newr.Borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            newr.Borders.Item[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            newr.Borders.Item[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            newr.Borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            newr.Borders.Item[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            newr.Borders.Item[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            newr.Rows.AutoFit();
            newr.Columns.AutoFit();
            newr.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        [ExcelCommand(MenuText = "BisectionSearch")]
        public static void BisectionSearch() {
            Range selection = App.Selection;
            if (selection == null) {
                MessageBox.Show("Invalid selection", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var range = selection.Resize[6, 2];

            var str1 = (string) range.Cells[1, 1].Value;
            var str2 = (string) range.Cells[2, 1].Value;
            var str3 = (string) range.Cells[3, 1].Value;
            var str4 = (string) range.Cells[4, 1].Value;
            var str5 = (string) range.Cells[5, 1].Value;
            var str6 = (string) range.Cells[6, 1].Value;

            Range xR = range.Cells[2, 2];
            Range yR = range.Cells[3, 2];
            double target = range.Cells[4, 2].Value;
            double lower = range.Cells[5, 2].Value;
            double upper = range.Cells[6, 2].Value;

            var f = yR.Formula;
            if (f == "formula") {
                xR.Value = "Error: y must be a formula";
                return;
            }

            if (upper < lower) {
                xR.Value = "Error: upper bound must be larger than lower bound";
                return;
            }
            xR.Value = upper;
            var vupper = yR.Value - target;
            if (vupper == 0) {
                xR.Value = upper;
                return;
            }
            xR.Value = lower;
            var vlower = yR.Value - target;
            if (vlower == 0) {
                xR.Value = lower;
                return;
            }
            if (vupper * vlower > 0) {
                xR.Value = "Error: bad range of x";
                return;
            }
            var guess = (lower + upper) * 0.5;
            while (upper - lower > Util.TOL) {
                xR.Value = guess;
                var vguess = yR.Value - target;
                if (vguess == 0) return;
                if (vguess * vupper < 0) {
                    lower = guess;
                }
                else {
                    upper = guess;
                    vupper = vguess;
                }
                guess = (lower + upper) * 0.5;
            }
            xR.Value = guess;
            //var values = (Array) selection.Cells.Value;
        }


        //private static object ViewFormulas([ExcelArgument(AllowReference = true)] object range)
        //{
        //    var theRef = (ExcelReference)range;
        //    var rows = theRef.RowLast - theRef.RowFirst + 1;
        //    var res = new object[rows, 1];
        //    for (var i = 0; i < rows; i++)
        //    {
        //        var cellRef = new ExcelReference(
        //            theRef.RowFirst + i, theRef.RowFirst + i,
        //            theRef.ColumnFirst, theRef.ColumnFirst,
        //            theRef.SheetId);
        //        var forluma = XlCall.Excel(XlCall.xlfGetFormula, cellRef).To<string>();
        //        if (!string.IsNullOrWhiteSpace(forluma))
        //        {
        //            var r = cellRef.ToRange();
        //            var f2 = App.ConvertFormula(forluma, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlA1, RelativeTo: r);
        //            var value = App.Evaluate(f2);
        //            res[i, 0] = f2;
        //        }
        //        else res[i, 0] = "No Formula";
        //    }
        //    return res;
        //}


        public static void WriteData()
        {
            var xlApp = (Application)ExcelDnaUtil.Application;

            var wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            var headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (var i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            var dataRange = ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
            xlApp.ActiveChart.SetSourceData(Source: dataRange);
        }
    }
}
