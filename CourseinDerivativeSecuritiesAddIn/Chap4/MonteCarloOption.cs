using ExcelDna.Integration;
using static System.Math;
using static CourseinDerivativeSecuritiesAddIn.Helper.NormalDist;

namespace CourseinDerivativeSecuritiesAddIn.Chap4 {
    public static class MonteCarloOption {
        [ExcelFunction("This outputs the row vector (call value, delta method 1, delta method 2)",
            Category = "Simulation")]
        public static object McEuroCall([ExcelArgument("initial stock price")] double s,
                                        [ExcelArgument("strike price")] double k,
                                        [ExcelArgument("risk-free rate")] double r,
                                        [ExcelArgument("volatility")] double σ,
                                        [ExcelArgument("dividend yield")] double q,
                                        [ExcelArgument("time to maturity")] double T,
                                        [ExcelArgument("number of simulations")] int m) {
            var logS0 = Log(s);
            var drift = (r - q - 0.5 * σ * σ) * T;
            var sigSqrtT = σ * Sqrt(T);
            var upChange = Log(1.01);
            var downChange = Log(0.99);
            var sumCall = 0D;
            var sumCallChange = 0D;
            var sumPathwise = 0D;
            double callV;
            for (var i = 1; i <= m; i++) {
                var logS = logS0 + drift + sigSqrtT * NormSample();
                callV = Max(0, Exp(logS) - k);
                sumCall += callV;
                var logSu = logS + upChange;
                var callVu = Max(0, Exp(logSu) - k);
                var logSd = logS + downChange;
                var callVd = Max(0, Exp(logSd) - k);
                sumCallChange += callVu - callVd;
                if (Exp(logS) > k)
                    sumPathwise += Exp(logS) / s;
            }
            callV = Exp(-r * T) * sumCall / m;
            var delta1 = Exp(-r * T) * sumCallChange / (m * 0.02 * s);
            var delta2 = Exp(-r * T) * sumPathwise / m;
            return new[] {callV, delta1, delta2};
        }

        [ExcelFunction("This outputs the row vector (call value, delta method 1, delta method 2)",
            Category = "Simulation")]
        public static object McEuroPut([ExcelArgument("initial stock price")] double s,
                                       [ExcelArgument("strike price")] double k,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility")] double σ,
                                       [ExcelArgument("dividend yield")] double q,
                                       [ExcelArgument("time to maturity")] double T,
                                       [ExcelArgument("number of simulations")] int m) {
            var logS0 = Log(s);
            var drift = (r - q - 0.5 * σ * σ) * T;
            var sigSqrtT = σ * Sqrt(T);
            var upChange = Log(1.01);
            var downChange = Log(0.99);
            var sumPut = 0D;
            var sumPutCHange = 0D;
            var sumPathwise = 0D;
            double putV;
            for (var i = 1; i <= m; i++) {
                var logS = logS0 + drift + sigSqrtT * NormSample();
                putV = Max(0, k - Exp(logS));
                sumPut += putV;
                var logSu = logS + upChange;
                var callVu = Max(0, k - Exp(logSu));
                var logSd = logS + downChange;
                var callVd = Max(0, k - Exp(logSd));
                sumPutCHange += callVu - callVd;
                if (Exp(logS) < k)
                    sumPathwise += Exp(logS) / s;
            }
            putV = Exp(-r * T) * sumPut / m;
            var delta1 = Exp(-r * T) * sumPutCHange / (m * 0.02 * s);
            var delta2 = Exp(-r * T) * sumPathwise / m;
            return new[] {putV, delta1, delta2};
        }

        [ExcelFunction("This returns the row vector (call value, standard error)", Category = "Simulation")]
        public static object McEuroCallGarch_1x2([ExcelArgument("initial stock price")] double s,
                                                 [ExcelArgument("strike price")] double k,
                                                 [ExcelArgument("risk-free rate")] double r,
                                                 [ExcelArgument("initial volatility")] double σ0,
                                                 [ExcelArgument("dividend yield")] double q,
                                                 [ExcelArgument("time to maturity")] double T,
                                                 [ExcelArgument("number of time periods")] int n,
                                                 [ExcelArgument("GARCH parameter")] double kappa,
                                                 [ExcelArgument("GARCH parameter")] double theta,
                                                 [ExcelArgument("GARCH parameter")] double lambda,
                                                 [ExcelArgument("number of simulations")] int m) {
            var dt = T / n;
            var sqrtdt = Sqrt(dt);
            var a = kappa * theta;
            var b = (1 - kappa) * lambda;
            var c = (1 - kappa) * (1 - lambda);
            var logS0 = Log(s);
            var sumCall = 0D;
            var sumCallSq = 0D;
            double callV;
            for (var i = 1; i <= m; i++) {
                var logS = logS0;
                var σ = σ0;
                for (var j = 1; j <= n; j++) {
                    var y = σ * NormSample();
                    logS += (r - q - 0.5 * σ * σ) * dt + sqrtdt * y;
                    σ = Sqrt(a + b * y * y + c * σ * σ);
                }
                callV = Max(0, Exp(logS) - k);
                sumCall += callV;
                sumCallSq += callV * callV;
            }
            callV = Exp(-r * T) * sumCall / m;
            var stdError = Exp(-r * T) * Sqrt((sumCallSq - sumCall * sumCall / m) / (m * (m - 1)));
            return new[] {callV, stdError};
        }

        [ExcelFunction("This returns the row vector (call value, standard error)", Category = "Simulation")]
        public static object McEuroPutGarch_1x2([ExcelArgument("initial stock price")] double s,
                                                [ExcelArgument("strike price")] double k,
                                                [ExcelArgument("risk-free rate")] double r,
                                                [ExcelArgument("initial volatility")] double σ0,
                                                [ExcelArgument("dividend yield")] double q,
                                                [ExcelArgument("time to maturity")] double T,
                                                [ExcelArgument("number of time periods")] int n,
                                                [ExcelArgument("GARCH parameter")] double kappa,
                                                [ExcelArgument("GARCH parameter")] double theta,
                                                [ExcelArgument("GARCH parameter")] double lambda,
                                                [ExcelArgument("number of simulations")] int m) {
            var dt = T / n;
            var sqrtdt = Sqrt(dt);
            var a = kappa * theta;
            var b = (1 - kappa) * lambda;
            var c = (1 - kappa) * (1 - lambda);
            var logS0 = Log(s);
            var sumPut = 0D;
            var sumPutSq = 0D;
            double putV;
            for (var i = 1; i <= m; i++) {
                var logS = logS0;
                var σ = σ0;
                for (var j = 1; j <= n; j++) {
                    var y = σ * NormSample();
                    logS += (r - q - 0.5 * σ * σ) * dt + sqrtdt * y;
                    σ = Sqrt(a + b * y * y + c * σ * σ);
                }
                putV = Max(0, k - Exp(logS));
                sumPut += putV;
                sumPutSq += putV * putV;
            }
            putV = Exp(-r * T) * sumPut / m;
            var stdError = Exp(-r * T) * Sqrt((sumPutSq - sumPut * sumPut / m) / (m * (m - 1)));
            return new[] {putV, stdError};
        }

    }
}
