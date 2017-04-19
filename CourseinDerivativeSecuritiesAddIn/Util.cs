using System;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn {
    public static class Util {
        internal static T To<T>(this object obj) => (T) obj;
        internal static double Pow(this double num, double pow = 2) => Math.Pow(num, pow);

        internal static object[,] ToColumn(this object[] input) {
            var l = input.Length;
            var output = new object[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        internal static object[,] ToColumn(this double[] input)
        {
            var l = input.Length;
            var output = new object[l, 1];
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
        public static object Sequence(double start=0, double step = 1) {
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
    }
}
