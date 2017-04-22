using ExcelDna.Integration;
using static System.Math;
using static DerivativeSecuritiesAddIn.Chap3.BlackScholes;
using static DerivativeSecuritiesAddIn.Chap7.BlackScholesExtension;
using static DerivativeSecuritiesAddIn.Helper.NormalDist;
using static DerivativeSecuritiesAddIn.Helper.BiNormalDist;
using static DerivativeSecuritiesAddIn.Util.Util;

namespace DerivativeSecuritiesAddIn.Chap8 {
    public static class ExoticOption {

        [ExcelFunction(Category = "Exotic Option")]
        public static double ForwardStartCall([ExcelArgument("initial stock price")] double s,
                                              [ExcelArgument("risk-free rate")] double r,
                                              [ExcelArgument("volatility")] double σ,
                                              [ExcelArgument("dividend yield")] double q,
                                              [ExcelArgument("time until the strike is set")] double tset,
                                              [ExcelArgument("time until call matures >= Tset")] double call) {
            var p1 = Exp(-q * call) * s;
            var p2 = Exp(-q * tset - r * (call - tset)) * s;
            return GenericOption(p1, p2, σ, call - tset);
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double ForwardStartPut([ExcelArgument("initial stock price")] double s,
                                             [ExcelArgument("risk-free rate")] double r,
                                             [ExcelArgument("volatility")] double σ,
                                             [ExcelArgument("dividend yield")] double q,
                                             [ExcelArgument("time until the strike is set")] double tset,
                                             [ExcelArgument("time until call matures >= Tset")] double call) {
            var p1 = Exp(-q * call) * s;
            var p2 = Exp(-q * tset - r * (call - tset)) * s;
            return GenericOption(p2, p1, σ, call - tset);
        }


        [ExcelFunction(Category = "Exotic Option")]
        public static double CallOnCall([ExcelArgument("initial stock price")] double s,
                                        [ExcelArgument("strike price of compound call")] double kc,
                                        [ExcelArgument("strike price of underlying call option")] double ku,
                                        [ExcelArgument("risk-free rate")] double r,
                                        [ExcelArgument("volatility")] double σ,
                                        [ExcelArgument("dividend yield")] double q,
                                        [ExcelArgument("time to maturity of compound call")] double tc,
                                        [ExcelArgument("time to maturity of underlying call >= Tc")] double tu) {
            var lower = 0D;
            var upper = Exp(q * (tu - tc)) * (kc + ku);
            var guess = 0.5 * lower + 0.5 * upper;
            var fupper = BsCall(upper, ku, r, σ, q, tu - tc) - kc;
            var fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
                }
            }
            var sstar = guess;
            var d1 = (Log(s / sstar) + (r - q + σ * σ / 2) * tc) / (σ * Sqrt(tc));
            var d2 = d1 - σ * Sqrt(tc);
            var d1Prime = (Log(s / ku) + (r - q + σ * σ / 2) * tu) / (σ * Sqrt(tu));
            var d2Prime = d1Prime - σ * Sqrt(tu);
            var rho = Sqrt(tc / tu);
            var n2 = NormDist(d2);
            var m1 = BiNormDist(d1, d1Prime, rho);
            var m2 = BiNormDist(d2, d2Prime, rho);
            return -Exp(-r * tc) * kc * n2 + Exp(-q * tu) * s * m1 - Exp(-r * tu) * ku * m2;
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double PutOnCall([ExcelArgument("initial stock price")] double s,
                                       [ExcelArgument("strike price of compound call")] double kc,
                                       [ExcelArgument("strike price of underlying call option")] double ku,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility")] double σ,
                                       [ExcelArgument("dividend yield")] double q,
                                       [ExcelArgument("time to maturity of compound call")] double tc,
                                       [ExcelArgument("time to maturity of underlying call >= Tc")] double tu) {
            var coc = CallOnCall(s, kc, ku, r, σ, q, tc, tu);
            var c = BsCall(s, ku, r, σ, q, tu);
            return coc + Exp(-r * tc) * kc - c;
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double CallOnPut([ExcelArgument("initial stock price")] double s,
                                       [ExcelArgument("strike price of compound call")] double kc,
                                       [ExcelArgument("strike price of underlying call option")] double ku,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility")] double σ,
                                       [ExcelArgument("dividend yield")] double q,
                                       [ExcelArgument("time to maturity of compound call")] double tc,
                                       [ExcelArgument("time to maturity of underlying call >= Tc")] double tu) {
            var lower = 0D;
            var upper = Exp(q * (tu - tc)) * (kc + ku);
            var guess = 0.5 * lower + 0.5 * upper;
            var fupper = BsCall(upper, ku, r, σ, q, tu - tc) - kc;
            var fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, ku, r, σ, q, tu - tc) - kc;
                }
            }
            var sstar = guess;
            var d1 = (Log(s / sstar) + (r - q + σ * σ / 2) * tc) / (σ * Sqrt(tc));
            var d2 = d1 - σ * Sqrt(tc);
            var d1Prime = (Log(s / ku) + (r - q + σ * σ / 2) * tu) / (σ * Sqrt(tu));
            var d2Prime = d1Prime - σ * Sqrt(tu);
            var rho = Sqrt(tc / tu);
            var n2 = NormDist(-d2);
            var m1 = BiNormDist(-d1, -d1Prime, rho);
            var m2 = BiNormDist(-d2, -d2Prime, rho);
            return -Exp(-r * tc) * kc * n2 + Exp(-r * tu) * ku * m2 - Exp(-q * tu) * s * m1;
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double PutOnPut([ExcelArgument("initial stock price")] double s,
                                      [ExcelArgument("strike price of compound call")] double kc,
                                      [ExcelArgument("strike price of underlying call option")] double ku,
                                      [ExcelArgument("risk-free rate")] double r,
                                      [ExcelArgument("volatility")] double σ,
                                      [ExcelArgument("dividend yield")] double q,
                                      [ExcelArgument("time to maturity of compound call")] double tc,
                                      [ExcelArgument("time to maturity of underlying call >= Tc")] double tu) {
            var cop = CallOnPut(s, kc, ku, r, σ, q, tc, tu);
            var p = BsPut(s, ku, r, σ, q, tu);
            return cop + Exp(-r * tc) * kc - p;
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double AmerCallDivd([ExcelArgument("initial stock price")] double s,
                                          [ExcelArgument("strike price")] double k,
                                          [ExcelArgument("risk-free rate")] double r,
                                          [ExcelArgument("volatility")] double σ,
                                          [ExcelArgument("cash dividend")] double div,
                                          [ExcelArgument("time until dividend payment")] double tDiv,
                                          [ExcelArgument("time until option matures >= TDiv")] double tCall) {
            var lessDiv = s - Exp(-r * tDiv) * div;
            if (div / k <= 1 - Exp(-r * (tCall - tDiv)))
                return BsCall(lessDiv, k, r, σ, 0, tCall);
            var upper = k;
            while (upper + div - k < BsCall(upper, k, r, σ, 0, tCall - tDiv))
                upper *= 2;
            var lower = 0D;
            var fupper = upper + div - k - BsCall(upper, k, r, σ, 0, tCall - tDiv);
            var guess = 0.5 * lower + 0.5 * upper;
            var fguess = guess + div - k - BsCall(guess, k, r, σ, 0, tCall - tDiv);
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = guess + div - k - BsCall(guess, k, r, σ, 0, tCall - tDiv);
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = guess + div - k - BsCall(guess, k, r, σ, 0, tCall - tDiv);
                }
            }
            var lessDivStar = guess;
            var d1 = (Log(lessDiv / lessDivStar) + (r + σ * σ / 2) * tDiv) / (σ * Sqrt(tDiv));
            var d2 = d1 - σ * Sqrt(tDiv);
            var d1Prime = (Log(lessDiv / k) + (r + σ * σ / 2) * tCall) / (σ * Sqrt(tCall));
            var d2Prime = d1Prime - σ * Sqrt(tCall);
            var rho = -Sqrt(tDiv / tCall);
            var n1 = NormDist(d1);
            var n2 = NormDist(d2);
            var m1 = BiNormDist(-d1, d1Prime, rho);
            var m2 = BiNormDist(-d2, d2Prime, rho);
            return lessDiv * n1 + Exp(-r * tDiv) * (div - k) * n2
                   + lessDiv * m1 - Exp(-r * tCall) * k * m2;
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double ChooserOption([ExcelArgument("initial stock price")] double s,
                                           [ExcelArgument("strike price of call option")] double kc,
                                           [ExcelArgument("strike price of put option")] double kp,
                                           [ExcelArgument("risk-free rate")] double r,
                                           [ExcelArgument("volatility")] double σ,
                                           [ExcelArgument("dividend yield")] double q,
                                           [ExcelArgument("time until choice must be made")] double T,
                                           [ExcelArgument("time until call matures >= T")] double tc,
                                           [ExcelArgument("time until put matures >= T")] double tp) {
            var lower = 0D;
            var upper = Exp(q * tc) * (kc + kp);
            var guess = 0.5 * kc + 0.5 * kp;
            var fupper = BsCall(upper, kc, r, σ, q, tc - T) - BsPut(upper, kp, r, σ, q, tp - T);
            var fguess = BsCall(guess, kc, r, σ, q, tc - T) - BsPut(guess, kp, r, σ, q, tp - T);
            while (upper - lower > TOL) {
                if (fupper * fguess < 0) {
                    lower = guess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, kc, r, σ, q, tc - T) - BsPut(guess, kp, r, σ, q, tp - T);
                }
                else {
                    upper = guess;
                    fupper = fguess;
                    guess = 0.5 * lower + 0.5 * upper;
                    fguess = BsCall(guess, kc, r, σ, q, tc - T) - BsPut(guess, kp, r, σ, q, tp - T);
                }
            }
            var sstar = guess;
            var d1 = (Log(s / sstar) + (r - q + σ.Pow() / 2) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var d1C = (Log(s / kc) + (r - q + σ.Pow() / 2) * tc) / (σ * Sqrt(tc));
            var d2C = d1C - σ * Sqrt(tc);
            var d1P = (Log(s / kp) + (r - q + σ.Pow() / 2) * tp) / (σ * Sqrt(tp));
            var d2P = d1P - σ * Sqrt(tp);
            var rhoc = Sqrt(T / tc);
            var rhop = Sqrt(T / tp);
            var m1C = BiNormDist(d1, d1C, rhoc);
            var m2C = BiNormDist(d2, d2C, rhoc);
            var m1P = BiNormDist(-d1, -d1P, rhop);
            var m2P = BiNormDist(-d2, -d2P, rhop);
            return Exp(-q * tc) * s * m1C - Exp(-r * tc) * kc * m2C + Exp(-r * tp) * kp * m2P - Exp(-q * tp) * s * m1P;
        }


        [ExcelFunction(Category = "Exotic Option")]
        public static double CallOnMax([ExcelArgument("price of stock 1")] double s1,
                                       [ExcelArgument("price of stock 2")] double s2,
                                       [ExcelArgument("strike price")] double k,
                                       [ExcelArgument("risk-free rate")] double r,
                                       [ExcelArgument("volatility of stock 1")] double sig1,
                                       [ExcelArgument("volatility of stock 2")] double sig2,
                                       [ExcelArgument("correlation")] double rho,
                                       [ExcelArgument("dividend yield of stock 1")] double q1,
                                       [ExcelArgument("dividend yield of stock 2")] double q2,
                                       [ExcelArgument("time to maturity")] double T) {
            var σ = Sqrt(sig2.Pow() - 2 * rho * sig1 * sig2 + sig1.Pow());
            var d1 = (Log(s1 / s2) + (q2 - q1 + σ.Pow() / 2) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var d11 = (Log(s1 / k) + (r - q1 + sig1.Pow() / 2) * T) / (sig1 * Sqrt(T));
            var d12 = d11 - sig1 * Sqrt(T);
            var d21 = (Log(s2 / k) + (r - q2 + sig2.Pow() / 2) * T) / (sig2 * Sqrt(T));
            var d22 = d21 - sig2 * Sqrt(T);
            var rho1 = (sig1 - rho * sig2) / σ;
            var rho2 = (sig2 - rho * sig1) / σ;
            var m1 = BiNormDist(d11, d1, rho1);
            var m2 = BiNormDist(d21, -d2, rho2);
            var m3 = BiNormDist(-d12, -d22, rho);
            return Exp(-q1 * T) * s1 * m1 + Exp(-q2 * T) * s2 * m2 + Exp(-r * T) * k * m3 - Exp(-r * T) * k;
        }


        [ExcelFunction(Category = "Exotic Option")]
        public static double DownAndOutCall([ExcelArgument("initial stock price")] double s,
                                            [ExcelArgument("strike price")] double k,
                                            [ExcelArgument("risk-free rate")] double r,
                                            [ExcelArgument("volatility")] double σ,
                                            [ExcelArgument("dividend yield")] double q,
                                            [ExcelArgument("time to maturity")] double T,
                                            [ExcelArgument("knock-out barrier < S")] double barrier) {
            double a, b;
            if (k > barrier) {
                a = s / k;
                b = barrier * barrier / (k * s);
            }
            else {
                a = s / barrier;
                b = barrier / s;
            }
            var d1 = (Log(a) + (r - q + 0.5 * σ.Pow()) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var d1Prime = (Log(b) + (r - q + 0.5 * σ.Pow()) * T) / (σ * Sqrt(T));
            var d2Prime = d1Prime - σ * Sqrt(T);
            var n1 = NormDist(d1);
            var n2 = NormDist(d2);
            var n1Prime = NormDist(d1Prime);
            var n2Prime = NormDist(d2Prime);
            var x = 1 + 2 * (r - q) / σ.Pow();
            var y = x - 2;
            var q1 = n1 - (barrier / s).Pow(x) * n1Prime;
            var q2 = n2 - (barrier / s).Pow(y) * n2Prime;
            return Exp(-q * T) * s * q1 - Exp(-r * T) * k * q2;
        }



        [ExcelFunction(Category = "Exotic Option")]
        public static double FloatingStrikeCall([ExcelArgument("initial stock price")] double s,
                                                  [ExcelArgument("risk-free rate")] double r,
                                                  [ExcelArgument("volatility")] double σ,
                                                  [ExcelArgument("dividend yield")] double q,
                                                  [ExcelArgument("time to maturity")] double T,
                                                  [ExcelArgument("minimum stock price during past life of contract")]
                                                  double sMin) {
            var d1 = (Log(s / sMin) + (r - q + 0.5 * σ.Pow()) * T) / (σ * Sqrt(T));
            var d2 = d1 - σ * Sqrt(T);
            var d2Prime = (Log(sMin / s) + (r - q - 0.5 * σ.Pow()) * T) / (σ * Sqrt(T));
            var n1 = NormDist(d1);
            var n2 = NormDist(d2);
            var n2Prime = NormDist(d2Prime);
            var x = 2 * (r - q) / σ.Pow();
            return Exp(-q * T) * s * n1 - Exp(-r * T) * sMin * n2
                   + 1 / x * (sMin / s).Pow(x) * Exp(-r * T) * sMin * n2Prime
                   - 1 / x * Exp(-q * T) * s * (1 - n1);
        }

        [ExcelFunction(Category = "Exotic Option")]
        public static double DiscreteGeomAvgPriceCall([ExcelArgument("initial stock price")] double s,
                                                              [ExcelArgument("stock price")] double k,
                                                              [ExcelArgument("risk-free rate")] double r,
                                                              [ExcelArgument("volatility")] double σ,
                                                              [ExcelArgument("dividend yield")] double q,
                                                              [ExcelArgument("time to maturity")] double T,
                                                              [ExcelArgument("number of time periods")] double n) {
            var dt = T / n;
            var nu = r - q - 0.5 * σ.Pow();
            var a = n * (n + 1) * (2 * n + 1) / 6;
            var v = Exp(-r * T) * s * Exp(((n + 1) * nu / 2 + σ.Pow() * a / (2 * n.Pow())) * dt);
            var sigavg = σ * Sqrt(a) / n.Pow(1.5);
            return BsCall(v, k, r, sigavg, 0, T);
        }
    }
}
