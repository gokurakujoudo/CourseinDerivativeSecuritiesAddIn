using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace DerivativeSecuritiesAddIn.Utility {
    public static class Util {
        private static readonly Application App = (Application)ExcelDnaUtil.Application;
        internal static T To<T>(this object obj) => (T) obj;
        internal static double Pow(this double num, double pow = 2) => Math.Pow(num, pow);

        public static Dictionary<TKey, TValue> ToDictionaryEx<TElement, TKey, TValue>(
            this IEnumerable<TElement> source, Func<TElement, TKey> keyGetter, Func<TElement, TValue> valueGetter) {
            var dict = new Dictionary<TKey, TValue>();
            foreach (var e in source) {
                var key = keyGetter(e);
                if (dict.ContainsKey(key))
                    continue;
                dict.Add(key, valueGetter(e));
            }
            return dict;
        }

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

        internal const double TOL = 1E-8;

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
            var range = App.Range[refText, Type.Missing];
            return range;
        }

        internal static string ToAddress(this Range range) {
            return range.Address[true, true, XlReferenceStyle.xlA1, true];
        }

        internal static Range ToRange(this string str) {
            try {
                var r = App.Range[str, Type.Missing];
                return r;
            }
            catch (Exception) {
                return null;
            }
        }

        [ExcelFunction(Category = "Utility", IsMacroType = true)]
        public static object Ref([ExcelArgument(AllowReference = true)] object range)
        {
            var rf = range.To<ExcelReference>();
            var refText = (string)XlCall.Excel(XlCall.xlfReftext, rf, true);
            return refText;
        }

        [ExcelFunction(Category = "Utility")]
        public static object Quote(string address) {
            var r = App.Range[address, Type.Missing];
            return r.Value;
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

        internal static Dictionary<string, object> ToDict(this object[,] value) {
            var n = value.GetLength(0);
            var dict = new Dictionary<string, object>();
            var mr = value.GetLowerBound(0);
            var mc = value.GetLowerBound(0);
            for (var i = mr; i < n + mr; i++) {
                var key = value[i, mc].ToString();
                var v = value[i, mc + 1];
                if (v is ExcelEmpty)
                    dict[key] = null;
                else if (v is int vi)
                    dict[key] = (double) vi;
                else if (v is string vs)
                    dict[key] = (object) vs.ToRange() ?? vs;
                else
                    dict[key] = v;
            }
            return dict;
        }

        internal static double ToDouble(this object obj) {
            if (obj is int i) return i;
            if (obj is double d) return d;
            throw new InvalidCastException();
        }

        internal static object ToExcelPrint(this object v) {
            if (v == null)
                return string.Empty;
            if (v is Range r)
                return r.ToAddress();
            if (v is int vi)
                return vi;
            if (v is double vd)
                return  vd;
            return  v.ToString();
        }
    }
}
