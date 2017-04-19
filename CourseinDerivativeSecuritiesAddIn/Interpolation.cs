using System;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn {
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


        [ExcelFunction(Category = "Interpolation")]
        public static object[,] CubicSpline(double[] xs, double[] ys, double[] newXs) {
            var n = xs.Length;
            var h = new double[n];
            var f = new double[n];
            var l = new double[n];
            var v = new double[n];
            var g = new double[n];

            for (var i = 0; i < n - 1; i++) {
                h[i] = xs[i + 1] - xs[i];
                f[i] = (ys[i + 1] - ys[i]) / h[i];
            }

            for (var i = 1; i < n - 1; i++) {
                l[i] = h[i] / (h[i - 1] + h[i]);
                v[i] = h[i - 1] / (h[i - 1] + h[i]);
                g[i] = 3 * (l[i] * f[i - 1] + v[i] * f[i]);
            }

            var b = new double[n];
            var tem = new double[n];
            var m = new double[n];
            var fn = (ys[n - 1] - ys[n - 2]) / (xs[n - 1] - xs[n - 2]);

            b[1] = v[1] / 2;
            for (var i = 2; i < n - 2; i++)
                b[i] = v[i] / (2 - b[i - 1] * l[i]);
            tem[1] = g[1] / 2;
            for (var i = 2; i < n - 1; i++)
                tem[i] = (g[i] - l[i] * tem[i - 1]) / (2 - l[i] * b[i - 1]);
            m[n - 2] = tem[n - 2];
            for (var i = n - 3; i > 0; i--)
                m[i] = tem[i] - b[i] * m[i + 1];
            m[0] = 3 * f[0] / 2.0;
            m[n - 1] = fn;
            var xlength = newXs.Length;
            var insertRes = new double[xlength];
            for (var i = 0; i < xlength; i++) {
                int j;
                for (j = 0; j < n; j++)
                    if (newXs[i] < xs[j])
                        break;
                j -= 1;
                if (j == -1 || j == n - 1) {
                    if (j == -1)
                        throw new Exception("Out of lower bound");
                    if (j == n - 1 && newXs[i] == xs[j])
                        insertRes[i] = ys[j];
                    else
                        throw new Exception("Out of upper bound");
                }
                else {
                    var p1 = (newXs[i] - xs[j + 1]) / (xs[j] - xs[j + 1]);
                    p1 = p1 * p1;
                    var p2 = (newXs[i] - xs[j]) / (xs[j + 1] - xs[j]);
                    p2 = p2 * p2;
                    var p3 = p1 * (1 + 2 * (newXs[i] - xs[j]) / (xs[j + 1] - xs[j])) * ys[j] + p2 *
                             (1 + 2 * (newXs[i] - xs[j + 1]) / (xs[j] - xs[j + 1])) * ys[j + 1];
                    var p4 = p1 * (newXs[i] - xs[j]) * m[j] + p2 * (newXs[i] - xs[j + 1]) * m[j + 1];
                    p4 = p4 + p3;
                    insertRes[i] = p4;
                }

            }
            return insertRes.ToColumn();
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object[,] CubicSplineFit(double[] xs, double[] ys, double newx) =>
            CubicSpline(xs, ys, new[] {newx});
    }
}
