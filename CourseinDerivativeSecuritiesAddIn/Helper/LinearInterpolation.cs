using System;
using System.Linq;
using DerivativeSecuritiesAddIn.Util;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Helper {
    public static class LinearInterpolation {

        [ExcelFunction(Category = "Interpolation")]
        public static double InLinearFit(double[] xs, double[] ys, double newX) {
            var lower = 0;
            var upper = xs.Length - 1;
            if (newX < xs[lower])
                upper = 1;
            else if (newX > xs[upper])
                lower = upper - 1;
            else {
                var mid = (lower + upper) / 2;
                while (lower <= mid && mid <= upper && upper - lower > 1) {
                    var x = xs[mid];
                    if (newX == x)
                        return ys[mid];
                    if (newX > x)
                        lower = mid;
                    if (newX < x)
                        upper = mid;
                    mid = (lower + upper) / 2;
                }
            }
            return ys[lower] + (newX - xs[lower]) / (xs[upper] - xs[lower]) * (ys[upper] - ys[lower]);
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object InLinearSeq(double lower, double upper) {
            var range = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            var count = range.RowLast - range.RowFirst + 1;
            var seq = LinearSeq_internal(lower, upper, count);
            if (seq == null)
                return "This is a Array Function";
            return seq.ToColumn();
        }

        internal static double[] LinearSeq_internal(double lower, double upper, int count) {
            var seq = new double[count];
            if (count == 1)
                return null;
            var delta = (upper - lower) / (count - 1);
            for (var i = 0; i < count; i++) {
                seq[i] = lower + i * delta;
            }
            return seq;
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object InLinearExpand(double[] xs, double[] ys) {
            var range = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            var count = range.RowLast - range.RowFirst + 1;
            if (count == 1)
                return "This is an array function";
            var newx = LinearSeq_internal(xs[0], xs[xs.Length - 1], count);
            var newy = newx.Select(x => InLinearFit(xs, ys, x)).ToArray();
            var result = new object[count, 2];
            for (var i = 0; i < count; i++) {
                result[i, 0] = newx[i];
                result[i, 1] = newy[i];
            }
            return result;
        }
    }
}
