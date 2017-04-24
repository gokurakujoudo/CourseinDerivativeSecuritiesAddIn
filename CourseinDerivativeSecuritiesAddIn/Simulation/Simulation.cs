using ExcelDna.Integration;
using static System.Math;
using static DerivativeSecuritiesAddIn.Chap3.BlackScholes;
using static DerivativeSecuritiesAddIn.Helper.NormalDist;
using static DerivativeSecuritiesAddIn.Utility.Util;

namespace DerivativeSecuritiesAddIn.Simulation {
    public static class Simulation {

        [ExcelFunction(Category = "Simulation")]
        public static object SimulatedDeltaHedgeProfit([ExcelArgument("initial stock price")] double s0,
                                                       [ExcelArgument("strike price")] double k,
                                                       [ExcelArgument("risk-free rate")] double r,
                                                       [ExcelArgument("volatility")] double sigma,
                                                       [ExcelArgument("dividend yield")] double q,
                                                       [ExcelArgument("time to maturity")] double T,
                                                       [ExcelArgument("expected rate of return")] double mu,
                                                       [ExcelArgument("number of time periods")] int n,
                                                       [ExcelArgument("number of simulations")] int m,
                                                       [ExcelArgument("percentile to be returned")] double pct) {
            var profit = new double[m];
            var dt = T / n;
            var sigSqrtdt = sigma * Sqrt(dt);
            var drift = (mu - q - 0.5 * sigma * sigma) * dt;
            var comp = Exp(r * dt);
            var div = Exp(q * dt) - 1;
            var logS0 = Log(s0);
            var call0 = BsCall(s0, k, r, sigma, q, T);
            var delta0 = BsCallDelta(s0, k, r, sigma, q, T);
            var cash0 = call0 - delta0 * s0;
            for (var i = 0; i < m; i++) {
                var logS = logS0;
                var cash = cash0;
                var s = s0;
                var delta = delta0;
                double newS;
                for (var j = 1; j < n; j++) {
                    logS += drift + sigSqrtdt * NormSample();
                    newS = Exp(logS);
                    var newDelta = BsCallDelta(newS, k, r, sigma, q, T - j * dt);
                    cash = comp * cash + delta * s * div - (newDelta - delta) * newS;
                    s = newS;
                    delta = newDelta;
                }
                logS += drift + sigSqrtdt * NormSample();
                newS = Exp(logS);
                var hedgeValue = comp * cash + delta * s * div + delta * newS;
                profit[i] = hedgeValue - Max(newS - k, 0);
            }
            return Percentile(profit, pct);
        }
    }
}
