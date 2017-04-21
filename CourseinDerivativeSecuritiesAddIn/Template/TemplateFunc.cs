using DerivativeSecuritiesAddIn.Util;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Template {
    public static class TemplateFunc {
        [ExcelFunction(Category = "Template")]
        public static object TempBsOption_6x1() =>
            new object[] {"S", "K", "r", "σ", "q", "T"}.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempBiOption_7x1() =>
            new object[] {"S", "K", "r", "σ", "q", "T", "N"}.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempMcOption_7x1() =>
            new object[] { "S", "K", "r", "σ", "q", "T", "M" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempMcGarch_4x1() =>
            new object[] {"N", "κ", "θ", "λ"}.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempBsOptionImpVol_6x1() =>
            new object[] {"S", "K", "r", "q", "T", "price"}.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempSiμlatedDelteHedge_10x1() =>
            new object[] {"S", "K", "r", "σ", "q", "T", "μ", "n", "m", "pct"}.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempGenericOption_4x1() =>
            new object[] { "P1", "P2", "σ", "T" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempExchangeOption_6x1() =>
            new object[] { "S1", "S2", "σ", "q1", "q2", "T" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempBlackOption_5x1() =>
            new object[] { "F", "K", "P", "σ", "T" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempDefferedExchangeOption_7x1() =>
            new object[] { "S1", "S2", "σ", "q1", "q2", "T", "T'" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempForwardStartOption_6x1() =>
            new object[] { "S", "r", "σ", "q", "T", "T'" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempCompoundOption_8x1() =>
            new object[] { "S", "K", "K'", "r", "σ", "q", "T", "T'" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempAmerCallDivi_7x1() =>
            new object[] { "S", "K", "r", "σ", "Div", "TDiv", "TCall" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempChooserOption_9x1() =>
            new object[] { "S", "Kc", "Kp", "r", "σ", "q", "T", "Tc", "Tp" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempCallOnMax_10x1() =>
            new object[] { "S1", "S2", "K", "r", "sig1", "sig2", "rho", "q1", "q2", "T" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempDownAndOutOption_7x1() =>
            new object[] { "S", "K", "r", "σ", "q", "T", "Barrier" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempFloatingStrikeOption_6x1() =>
            new object[] { "S", "r", "σ", "q", "T", "SMin" }.ToColumn();

        [ExcelFunction(Category = "Template")]
        public static object TempDisGeomAvegPriceOption_7x1() =>
            new object[] { "S", "K", "r", "σ", "q", "T", "N" }.ToColumn();

    }
}
