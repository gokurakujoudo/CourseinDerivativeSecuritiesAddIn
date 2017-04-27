using System.Linq;
using System.Windows;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.ObjectSystem
{
    public static class SharpObjectExcelTemplate {
        private static readonly Application App = (Application) ExcelDnaUtil.Application;

        internal static void CreateObjTemp(SharpObjectTemplate temp) {
            Range selection = App.Selection;
            if (selection == null) {
                MessageBox.Show("Invalid selection", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var propArray = temp.PropertyList.ToArray();
            var cpropArray = temp.ComputedPropertyList.ToArray();
            var nprop = propArray.Length;
            var ncporp = cpropArray.Length;
            var range = selection.Resize[2 + nprop + ncporp + 1, 2];

            var rangeIn = selection.Resize[2 + nprop, 2];
            var cellCall = range.Cells[2 + nprop + 1, 2];
            var rangeCall = range.Cells[2 + nprop + 1, 1].Resize[1, 2];
            var rangeOut = range.Cells[2 + nprop + 2, 1].Resize[ncporp, 2];

            var valueIn = new object[2 + nprop, 2];
            valueIn[0, 0] = SharpObjectHelper.ID;
            valueIn[1, 0] = SharpObjectHelper.TYPE;
            valueIn[1, 1] = temp.TypeName;
            for (var i = 0; i < nprop; i++) {
                var propName = propArray[i].Key;
                var propDef = propArray[i].Value;
                valueIn[i + 2, 0] = propName;
                if (temp.PropertyDescription.TryGetValue(propName, out var propDes)) {
                    if (propDef != null)
                        valueIn[i + 2, 1] = $"[{propDes}] = {propDef.ToExcelPrint()}";
                    else
                        valueIn[i + 2, 1] = $"{propDes}";
                }
                else {
                    if (propDef != null)
                        valueIn[i + 2, 1] = $"[{propDef.ToExcelPrint()}]";
                    else
                        valueIn[i + 2, 1] = "NO DEFINATION";
                }
            }
            rangeIn.Value = valueIn;
            rangeCall.Value = new object[] {"OBJECT", null};
            cellCall.Formula = $"={nameof(SharpObjectHelper.CreateObj)}(R[-{nprop + 2}]C[-1]:R[-1]C)";
            rangeOut.FormulaArray = $"={nameof(SharpObjectHelper.ViewObjComputedProp)}(R[-1]C[1])";

            range.Cells[1, 1].Resize[1, 2].Interior.Color = XlRgbColor.rgbSkyBlue;
            rangeCall.Interior.Color = XlRgbColor.rgbLightGray;
            rangeOut.Interior.Color = XlRgbColor.rgbLightGray;

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
