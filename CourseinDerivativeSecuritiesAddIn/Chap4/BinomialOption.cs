using ExcelDna.Integration;
using static System.Math;

namespace DerivativeSecuritiesAddIn.Chap4
{
    public static class BinomialOption {
        [ExcelFunction(Category = "Binomial")]
        public static double BiEuroCall([ExcelArgument("initial stock price")] double s,
                                        [ExcelArgument("strike price")] double k,
                                        [ExcelArgument("risk-free rate")] double r,
                                        [ExcelArgument("volatility")] double σ,
                                        [ExcelArgument("dividend yield")] double q,
                                        [ExcelArgument("time to maturity")] double T,
                                        [ExcelArgument("number of time periods")] int n) {

            var dt = T / n;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var pd = 1 - pu;
            var u2 = u * u;
            s *= Pow(d, n);
            var prob = Pow(pd, n);
            var callV = prob * Max(s - k, 0);
            for (var i = 1; i <= n; i++) {
                s *= u2;
                prob *= pu / pd * (n - i + 1) / i;
                callV += prob * Max(s - k, 0);
            }
            return Exp(-r * T) * callV;
        }

        [ExcelFunction(Category = "Binomial")]
        public static double BiEuroPut([ExcelArgument("initial stock price")] double s,
                                       [ExcelArgument("strike price")] double k,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility")] double σ,
                                       [ExcelArgument("dividend yield")] double q,
                                       [ExcelArgument("time to maturity")] double T,
                                       [ExcelArgument("number of time periods")] int n) {

            var dt = T / n;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var pd = 1 - pu;
            var u2 = u * u;
            s *= Pow(d, n);
            var prob = Pow(pd, n);
            var putV = prob * Max(s - k, 0);
            for (var i = 1; i <= n; i++) {
                s *= u2;
                prob *= pu / pd * (n - i + 1) / i;
                putV += prob * Max(k - s, 0);
            }
            return Exp(-r * T) * putV;
        }

        [ExcelFunction(Category = "Binomial")]
        public static double BiAmerPut([ExcelArgument("initial stock price")] double s,
                                       [ExcelArgument("strike price")] double k,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility")] double σ,
                                       [ExcelArgument("dividend yield")] double q,
                                       [ExcelArgument("time to maturity")] double T,
                                       [ExcelArgument("number of time periods")] int n) {
            var putV = new double[n + 1];
            var dt = T / n;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var dpu = Exp(-r * dt) * pu;
            var dpd = Exp(-r * dt) * (1 - pu);
            var u2 = u * u;
            var st = s * Pow(d, n);
            putV[0] = Max(k - st, 0);
            for (var j = 1; j <= n; j++) {
                st *= u2;
                putV[j] = Max(k - st, 0);
            }
            for (var i = n - 1; i >= 0; i--) {
                st = s * Pow(d, i);
                putV[0] = Max(k - st, dpd * putV[0] + dpu * putV[1]);
                for (var j = 1; j <= i; j++) {
                    st *= u2;
                    putV[j] = Max(k - st, dpd * putV[j] + dpu * putV[j + 1]);
                }
            }
            return putV[0];
        }

        [ExcelFunction(Category = "Binomial")]
        public static double BiAmerCall([ExcelArgument("initial stock price")] double s,
                                        [ExcelArgument("strike price")] double k,
                                        [ExcelArgument("risk-free rate")] double r,
                                        [ExcelArgument("volatility")] double σ,
                                        [ExcelArgument("dividend yield")] double q,
                                        [ExcelArgument("time to maturity")] double T,
                                        [ExcelArgument("number of time periods")] int n) {
            var callV = new double[n + 1];
            var dt = T / n;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var dpu = Exp(-r * dt) * pu;
            var dpd = Exp(-r * dt) * (1 - pu);
            var u2 = u * u;
            var st = s * Pow(d, n);
            callV[0] = Max(st - k, 0);
            for (var j = 1; j <= n; j++) {
                st *= u2;
                callV[j] = Max(st - k, 0);
            }
            for (var i = n - 1; i >= 0; i--) {
                st = s * Pow(d, i);
                callV[0] = Max(st - k, dpd * callV[0] + dpu * callV[1]);
                for (var j = 1; j <= i; j++) {
                    st *= u2;
                    callV[j] = Max(st - k, dpd * callV[j] + dpu * callV[j + 1]);
                }
            }
            return callV[0];
        }

        [ExcelFunction("This returns the row vector (put value, delta, gamma)", Category = "Binomial")]
        public static object BiAmerPutDG_1x3([ExcelArgument("initial stock price")] double s,
                                             [ExcelArgument("strike price")] double k,
                                             [ExcelArgument("risk-free rate")] double r,
                                             [ExcelArgument("volatility")] double σ,
                                             [ExcelArgument("dividend yield")] double q,
                                             [ExcelArgument("time to maturity")] double T,
                                             [ExcelArgument("number of time periods")] int n) {
            var putV = new double[n + 3];
            var dt = T / n;
            var newN = n + 2;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var dpu = Exp(-r * dt) * pu;
            var dpd = Exp(-r * dt) * (1 - pu);
            var u2 = u * u;
            var st = s * Pow(d, newN);
            putV[0] = Max(k - st, 0);
            for (var j = 1; j <= newN; j++) {
                st *= u2;
                putV[j] = Max(k - st, 0);
            }
            for (var i = newN - 1; i >= 2; i--) {
                st = s * Pow(d, i);
                putV[0] = Max(k - st, dpd * putV[0] + dpu * putV[1]);
                for (var j = 1; j <= i; j++) {
                    st *= u2;
                    putV[j] = Max(k - st, dpd * putV[j] + dpu * putV[j + 1]);
                }
            }
            var su = s * u2;
            var sd = s / u2;
            var deltau = (putV[2] - putV[1]) / (su - s);
            var deltad = (putV[1] - putV[0]) / (s - sd);
            var distance = s * (u2 - d * d);
            var delta = (putV[2] - putV[0]) / distance;
            var gamma = 2 * (deltau - deltad) / distance;
            return new[] {putV[1], delta, gamma};
        }

        [ExcelFunction("This returns the row vector (put value, delta, gamma)", Category = "Binomial")]
        public static object BiAmerCallDG_1x3([ExcelArgument("initial stock price")] double s,
                                              [ExcelArgument("strike price")] double k,
                                              [ExcelArgument("risk-free rate")] double r,
                                              [ExcelArgument("volatility")] double σ,
                                              [ExcelArgument("dividend yield")] double q,
                                              [ExcelArgument("time to maturity")] double T,
                                              [ExcelArgument("number of time periods")] int n) {
            var callV = new double[n + 3];
            var dt = T / n;
            var newN = n + 2;
            var u = Exp(σ * Sqrt(dt));
            var d = 1 / u;
            var pu = (Exp((r - q) * dt) - d) / (u - d);
            var dpu = Exp(-r * dt) * pu;
            var dpd = Exp(-r * dt) * (1 - pu);
            var u2 = u * u;
            var st = s * Pow(d, newN);
            callV[0] = Max(st - k, 0);
            for (var j = 1; j <= newN; j++) {
                st *= u2;
                callV[j] = Max(st - k, 0);
            }
            for (var i = newN - 1; i >= 2; i--) {
                st = s * Pow(d, i);
                callV[0] = Max(st - k, dpd * callV[0] + dpu * callV[1]);
                for (var j = 1; j <= i; j++) {
                    st *= u2;
                    callV[j] = Max(st - k, dpd * callV[j] + dpu * callV[j + 1]);
                }
            }
            var su = s * u2;
            var sd = s / u2;
            var deltau = (callV[2] - callV[1]) / (su - s);
            var deltad = (callV[1] - callV[0]) / (s - sd);
            var distance = s * (u2 - d * d);
            var delta = (callV[2] - callV[0]) / distance;
            var gamma = 2 * (deltau - deltad) / distance;
            return new[] {callV[1], delta, gamma};
        }
    }
}
