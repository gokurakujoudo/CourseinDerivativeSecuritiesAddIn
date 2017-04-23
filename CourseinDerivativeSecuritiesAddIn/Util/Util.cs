using System;
using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace DerivativeSecuritiesAddIn.Util {
    public static class Util {
        internal static Application app = (Application)ExcelDnaUtil.Application;
        internal static T To<T>(this object obj) => (T) obj;
        internal static double Pow(this double num, double pow = 2) => Math.Pow(num, pow);

        internal static object[,] ToColumn(this object[] input) {
            var l = input.Length;
            var output = new object[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        internal static object[,] ToColumn(this double[] input) {
            var l = input.Length;
            var output = new object[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        internal static T[,] ToColumn<T>(this T[] input) {
            var l = input.Length;
            var output = new T[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        internal const double TOL = 1E-6;

        internal static double Percentile(double[] sequence, double excelPercentile) {
            Array.Sort(sequence);
            var length = sequence.Length;
            var n = (length - 1) * excelPercentile + 1;
            if (n == 1d) return sequence[0];
            if (n == length) return sequence[length - 1];
            var k = (int) n;
            var d = n - k;
            return sequence[k - 1] + d * (sequence[k] - sequence[k - 1]);
        }

        [ExcelFunction("Generate sequence", Category = "Template")]
        public static object Sequence(double start = 0, double step = 1) {
            var range = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            if (range.ColumnFirst == range.ColumnLast) {
                var count = range.RowLast - range.RowFirst + 1;
                var result = new object[count];
                for (var i = 0; i < count; i++)
                    result[i] = start + i * step;
                return result.ToColumn();
            }
            if (range.RowFirst == range.RowLast) {
                var count = range.ColumnLast - range.ColumnFirst + 1;
                var result = new object[count];
                for (var i = 0; i < count; i++)
                    result[i] = start + i * step;
                return result;
            }
            return "Must be a row or a column";
        }

        internal static double Div(this double d1, double d2) =>
            d2 == 0 ? 0 : d1 / d2;

        [ExcelFunction(Category = "Utility")]
        public static object About() => "Created by GitHub:Gokurakujoudu 2017";

        internal static Range ToRange(this ExcelReference xlref) {
            var refText = (string) XlCall.Excel(XlCall.xlfReftext, xlref, true);
            var range = app.Range[refText, Type.Missing];
            return range;
        }

        [ExcelFunction(Category = "Utility", IsMacroType = true)]
        public static object ToAddress([ExcelArgument(AllowReference = true)] object range) {
            var rf = range.To<ExcelReference>();
            var refText = (string) XlCall.Excel(XlCall.xlfReftext, rf, true);
            return refText;
        }

        [ExcelFunction(Category = "Utility", IsMacroType = true)]
        public static object ViewFormulas([ExcelArgument(AllowReference = true)] object range) {
            var theRef = (ExcelReference) range;
            var rows = theRef.RowLast - theRef.RowFirst + 1;
            var res = new object[rows, 1];
            for (var i = 0; i < rows; i++) {
                var cellRef = new ExcelReference(
                    theRef.RowFirst + i, theRef.RowFirst + i,
                    theRef.ColumnFirst, theRef.ColumnFirst,
                    theRef.SheetId);
                var forluma = XlCall.Excel(XlCall.xlfGetFormula, cellRef).To<string>();
                if (!string.IsNullOrWhiteSpace(forluma)) {
                    var r = cellRef.ToRange();
                    var f2 = app.ConvertFormula(forluma, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlA1, RelativeTo: r);
                    var value = app.Evaluate(f2);
                    res[i, 0] = f2;
                }
                else res[i, 0] = "No Formula";
            }
            return res;
        }

        [ExcelFunction(Category = "Utility")]
        public static object AutoIndex(object[] reference) {
            var index = 1;
            var n = reference.Length;
            var auto = new object[n];
            for (var i = 0; i < auto.Length; i++) {
                if (reference[i] is ExcelEmpty) {
                    auto[i] = "";
                    continue;
                }
                auto[i] = index;
                index++;
            }
            return auto.ToColumn();
        }
    }
}
