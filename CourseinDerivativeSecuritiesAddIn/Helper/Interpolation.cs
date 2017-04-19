using System;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Helper {
    public static class Interpolation {

        [ExcelFunction(Category = "Interpolation")]
        public static double LinearFit(double[] xs, double[] ys, double newX) {
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
    }
}
