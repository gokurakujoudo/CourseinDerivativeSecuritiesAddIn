using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using DerivativeSecuritiesAddIn.Chap3;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.Utility {
    public static class TemplateSubBase {
        private const string TYPE = "TYPE";

        private static readonly Application App = (Application) ExcelDnaUtil.Application;

        private static readonly Dictionary<string, MethodInfo> Funcdict = Assembly
            .GetExecutingAssembly()
            .GetTypes()
            .Where(type => type.Namespace != null && type.Namespace.Contains("DerivativeSecuritiesAddIn"))
            .SelectMany(type => type.GetMethods())
            .ToDictionaryEx(m => m.Name.ToUpper(), m => m);

        public static void TestTemp() => CreateTemp(typeof(BlackScholes).GetMethod(nameof(BlackScholes.BsCallDelta)));

        internal static void CreateTemp(MethodInfo method) {
            Range selection = App.Selection;
            if (selection == null) {
                MessageBox.Show("Invalid selection", "", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var refs = method.GetParaInfo();
            var n = refs.Length;
            var range = selection.Resize[n + 2, 2];
            var value = new object[n + 2, 2];
            value[0, 0] = TYPE;
            value[0, 1] = method.Name.ToUpper();
            for (var i = 0; i < n; i++) {
                value[i + 1, 0] = refs[i].Name;
                value[i + 1, 1] = refs[i].Optional ? $"[{refs[i].Type.Name}] = {refs[i].Default}" : refs[i].Type.Name;
            }
            value[n + 1, 0] = "RESULT";
            range.Value = value;
            range.Cells[n + 2, 2].FormulaR1C1 = $"={nameof(TempCall)}(R[-{n + 1}]C[-1]:R[-1]C)";

            range.Cells[1, 1].Interior.Color = XlRgbColor.rgbSkyBlue;
            range.Cells[1, 2].Interior.Color = XlRgbColor.rgbSkyBlue;
            range.Cells[n + 2, 1].Interior.Color = XlRgbColor.rgbLightGray;
            range.Cells[n + 2, 2].Interior.Color = XlRgbColor.rgbLightGray;

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

        [ExcelFunction(Category = "Utility")]
        public static object TempCall(object[,] input) {
            if (input == null) return "Error Range";
            var dict = new Dictionary<string, object>();
            for (var i = 0; i < input.GetLength(0); i++)
                dict[input[i, 0].To<string>()] = input[i, 1];
            var methodName = dict[TYPE].ToString();
            if (string.IsNullOrEmpty(methodName) || !Funcdict.ContainsKey(methodName)) return "Error Method";
            var method = Funcdict[methodName];
            var paraInfo = method.GetParaInfo();
            var n = paraInfo.Length;
            var paras = new object[n];
            for (var i = 0; i < n; i++) {
                var paraName = paraInfo[i].Name;
                if (dict.ContainsKey(paraName) && !(dict[paraName] is ExcelEmpty)) {
                    var raw = dict[paraName];
                    if (raw is double d && paraInfo[i].Type == typeof(int))
                        paras[i] = (int) d;
                    else if (paraInfo[i].Type == typeof(string))
                        paras[i] = raw.To<string>();
                    else paras[i] = raw;
                }
                else if (paraInfo[i].Optional)
                    paras[i] = paraInfo[i].Default;
                else return $"Invalid value in {paraName}";
            }
            return method.Invoke(null, paras);
        }
    }


    public class ExNameAttribute : Attribute {
        public ExNameAttribute(string name) { this.Name = name; }
        public string Name { get; }
    }
}
