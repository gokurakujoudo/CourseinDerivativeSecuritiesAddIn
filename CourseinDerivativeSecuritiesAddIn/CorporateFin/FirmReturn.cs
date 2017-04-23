using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.CorporateFin
{
    public static class FirmReturn {
        [ExcelFunction(Category = "Corporate Finance")]
        public static object CorpReturnOnAssetConstDOverV(double re, double rd, double dOverV, double tax) =>
            (1 - dOverV) * re + dOverV * rd * (1 - tax);

        [ExcelFunction(Category = "Corporate Finance")]
        public static object CorpReturnOnAssetConstD(double re, double rd, double e, double d, double tax) {
            var dt = d * (1 - tax);
            return e / (e + dt) * re + dt / (e + dt) * rd;
        }

        [ExcelFunction(Category = "Corporate Finance")]
        public static object CorpReturnOnEquityConstDOverV(double ra, double rd, double dOverV, double tax) {
            var e = 1 - dOverV;
            return ra / e - dOverV / e * rd * (1 - tax);
        }
            

        [ExcelFunction(Category = "Corporate Finance")]
        public static object CorpReturnOnEquityConstD(double ra, double rd, double e, double d, double tax)
        {
            var dt = d * (1 - tax);
            return ra + dt / e * (ra - rd);
        }
    }
}
