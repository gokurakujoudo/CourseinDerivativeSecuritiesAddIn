using System.Linq;
using System.Windows;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.Alpha
{
    public static class SharpObjectExcelTemplate
    {
        private static readonly Application App = (Application)ExcelDnaUtil.Application;

        internal static void CreateObjTemp(SharpObjectTemplate temp) {
            Range selection = App.Selection;
            if (selection == null)
            {
                MessageBox.Show("Invalid selection", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var propArray = temp.PropertyList.ToArray();
            var cpropArray = temp.ComputedPropertyList.ToArray();
            var nprop = propArray.Length;
            var ncporp = cpropArray.Length;
            var range = selection.Resize[2 + nprop + ncporp + 1, 2];

            var rangeIn = selection.Resize[2 + nprop, 2];
            var rangeOut = range.Cells[2 + nprop + 1, 1].Resize[ncporp + 1, 2];
            var cellCall = range.Cells[2 + nprop + 1, 2];

            var valueIn = new object[2 + nprop, 2];
            valueIn[0, 0] = SharpObjectHelper.ID;
            valueIn[1, 0] = SharpObjectHelper.TYPE;
            valueIn[1, 1] = temp.TypeName;
            for (var i = 0; i < nprop; i++) {
                var propName = propArray[i].Key;
                var propDef = propArray[i].Value;
                var propDes = temp.PropertyDescription[propName];
            }
            rangeIn.Value = valueIn;//TODO

            //var value = new object[n + count + 1, 2];

            //value[n + 1, 0] = "RESULT";
            //range.Value = value;
            //range.Cells[n + 2, 2].Resize[count, 1].FormulaArray = $"={nameof(TempCall)}(R[-{n + 1}]C[-1]:R[-1]C)";
            //range.Cells[1, 1].Resize[1, 2].Interior.Color = XlRgbColor.rgbSkyBlue;
            //range.Cells[n + 2, 1].Resize[count, 2].Interior.Color = XlRgbColor.rgbLightGray;



            range.Borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Item[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            range.Rows.AutoFit();
            range.Columns.AutoFit();
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }
    }
}
