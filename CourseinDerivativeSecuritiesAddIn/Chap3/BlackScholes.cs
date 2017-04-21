using ExcelDna.Integration;
using static System.Math;
using static DerivativeSecuritiesAddIn.Helper.NormalDist;
using static DerivativeSecuritiesAddIn.Util.Util;

namespace DerivativeSecuritiesAddIn.Chap3 {
    public static class BlackScholes {

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCall(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + T * (r - q + σ * σ / 2)) / (σ * Sqrt(T));
            var d2 = (Log(s / k) + T * (r - q - σ * σ / 2)) / (σ * Sqrt(T));
            var n1 = NormDist(d1);
            var n2 = NormDist(d2);
            return Exp(-q * T) * s * n1 - Exp(-r * T) * k * n2;
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCallDelta(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + T * (r - q + σ * σ / 2)) / (σ * Sqrt(T));
            var n1 = NormDist(d1);
            return Exp(-q * T) * n1;
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCallGamma(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + (r - q + 0.5 * σ * σ) * T) / (σ * Sqrt(T));
            var nd1 = Exp(-d1 * d1 / 2) / Sqrt(2 * PI);
            return Exp(-q * T) * nd1 / (s * σ * Sqrt(T));
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCallTheta(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + (r - q + 0.5 * σ * σ) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var n1 = NormDist(d1);
            var n2 = NormDist(d2);
            var nd1 = Exp(-d1 * d1 / 2) / Sqrt(2 * PI);
            return -Exp(-q * T) * s * nd1 * σ / 2 / Sqrt(T) + q * Exp(-q * T) * s * n1 - r * Exp(-r * T) * k * n2;
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCallVega(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + (r - q + 0.5 * σ * σ) * T) / (σ * Sqrt(T));
            var nd1 = Exp(-d1 * d1 / 2) / Sqrt(2 * PI);
            return Exp(-q * T) * s * nd1 * Sqrt(T);
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsCallRho(double s, double k, double r, double σ, double q, double T) {
            var d1 = (Log(s / k) + (r - q + 0.5 * σ * σ) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var n2 = NormDist(d2);
            return T * Exp(-r * T) * k * n2;
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPut(double s, double k, double r, double σ, double q, double T) =>
            BsCall(s, k, r, σ, q, T) + Exp(-r * T) * k - Exp(-q * T) * s;

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPutDelta(double s, double k, double r, double σ, double q, double T) =>
            BsCallDelta(s, k, r, σ, q, T) - Exp(-q * T);

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPutGamma(double s, double k, double r, double σ, double q, double T) =>
            BsCallGamma(s, k, r, σ, q, T);

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPutTheta(double s, double k, double r, double σ, double q, double T) =>
            BsCallTheta(s, k, r, σ, q, T) - r * Exp(-r * T) * k + q * Exp(-q * T) * s;

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPutVega(double s, double k, double r, double σ, double q, double T) =>
            BsCallVega(s, k, r, σ, q, T);

        [ExcelFunction(Category = "BlackScholes")]
        public static double BsPutRho(double s, double k, double r, double σ, double q, double T) =>
            BsCallRho(s, k, r, σ, q, T) - T * Exp(-r * T) * k;

        [ExcelFunction(Category = "BlackScholes")]
        public static object BsCallImpVol(double s, double k, double r, double q, double T, double callPrice) {
            if (callPrice < Exp(-q * T) * s - Exp(-r * T) * k)
                return "Option price violates the arbitrage bound.";
            var lower = 0D;
            var upper = 1D;
            var fupper = BsCall(s, k, r, upper, q, T) - callPrice;
            while (fupper < 0) {
                upper *= 2;
                fupper = BsCall(s, k, r, upper, q, T) - callPrice;
            }
            var guess = 0.5 * lower + 0.5 * upper;
            var fguess = BsCall(s, k, r, guess, q, T) - callPrice;
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(s, k, r, guess, q, T) - callPrice;
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(s, k, r, guess, q, T) - callPrice;
                }
            }
            return guess;
        }

        [ExcelFunction(Category = "BlackScholes")]
        public static object BsPutImpVol(double s, double k, double r, double q, double T, double putPrice) {
            if (putPrice < Exp(-r * T) * k - Exp(-q * T) * s)
                return "Option price violates the arbitrage bound.";
            var lower = 0D;
            var upper = 1D;
            var fupper = BsPut(s, k, r, upper, q, T) - putPrice;
            while (fupper < 0) {
                upper *= 2;
                fupper = BsPut(s, k, r, upper, q, T) - putPrice;
            }
            var guess = 0.5 * lower + 0.5 * upper;
            var fguess = BsPut(s, k, r, guess, q, T) - putPrice;
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsPut(s, k, r, guess, q, T) - putPrice;
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsPut(s, k, r, guess, q, T) - putPrice;
                }
            }
            return guess;
        }
    }
}
