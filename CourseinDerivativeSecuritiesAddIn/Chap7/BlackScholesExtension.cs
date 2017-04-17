using ExcelDna.Integration;
using static System.Math;
using static DerivativeSecuritiesAddIn.Helper.NormalDist;

namespace DerivativeSecuritiesAddIn.Chap7 {
    public static class BlackScholesExtension {

        [ExcelFunction(Category = "Generic Option")]
        public static double GenericOption([ExcelArgument("present value of asset to be received")] double p1,
                                           [ExcelArgument("present value of asset to be delivered")] double p2,
                                           [ExcelArgument("volatility")] double σ,
                                           [ExcelArgument("time to maturity")] double T) {
            var x = (Log(p1 / p2) + 0.5 * σ * σ * T) / (σ * Sqrt(T));
            var y = x - σ * Sqrt(T);
            var n1 = NormDist(x);
            var n2 = NormDist(y);
            return p1 * n1 - p2 * n2;
        }

        [ExcelFunction(Category = "Generic Option")]
        public static double ExchangeOption([ExcelArgument("price of asset to be received")] double s1,
                                            [ExcelArgument("price of asset to be delivered")] double s2,
                                            [ExcelArgument("volatility of ratio of prices")] double σ,
                                            [ExcelArgument("dividend yield of asset to be received")] double q1,
                                            [ExcelArgument("dividend yield of asset to be delivered")] double q2,
                                            [ExcelArgument("time to maturity")] double T) =>
            GenericOption(Exp(-q1 * T) * s1, Exp(-q2 * T) * s2, σ, T);


        [ExcelFunction("To value a futures option, input F = futures price and" +
                       " P = price of discount bond maturing when option matures",
            Category = "Generic Option")]
        public static double BlackCall([ExcelArgument("forward price")] double f,
                                       [ExcelArgument("strike price")] double k,
                                       [ExcelArgument("price of discount bond maturing when forward matures")] double p,
                                       [ExcelArgument("volatility of forward price")] double σ,
                                       [ExcelArgument("time to maturity")] double T) =>
            GenericOption(p * f, p * k, σ, T);

        [ExcelFunction("To value a futures option, input F = futures price and" +
                       " P = price of discount bond maturing when option matures",
            Category = "Generic Option")]
        public static double BlackPut([ExcelArgument("forward price")] double f,
                                      [ExcelArgument("strike price")] double k,
                                      [ExcelArgument("price of discount bond maturing when forward matures")] double p,
                                      [ExcelArgument("volatility of forward price")] double σ,
                                      [ExcelArgument("time to maturity")] double T) =>
            GenericOption(p * k, p * f, σ, T);


        [ExcelFunction(Category = "Generic Option")]
        public static double DeferredExchangeOption([ExcelArgument("price of asset to be received")] double s1,
                                                    [ExcelArgument("price of asset to be delivered")] double s2,
                                                    [ExcelArgument("volatility of ratio of prices")] double σ,
                                                    [ExcelArgument("dividend yield of asset to be received")] double q1,
                                                    [ExcelArgument("dividend yield of asset to be delivered")]
                                                    double q2,
                                                    [ExcelArgument("time to maturity of option")] double tmat,
                                                    [ExcelArgument("time until exchange >= TOption")] double texch) =>
            GenericOption(Exp(-q1 * texch) * s1, Exp(-q2 * texch) * s2, σ, tmat);

    }
}
