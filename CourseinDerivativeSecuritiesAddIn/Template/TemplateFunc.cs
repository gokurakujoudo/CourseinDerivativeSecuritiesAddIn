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
        public static object TempSimulatedDelteHedge_10x1() =>
            new object[] {"S", "K", "r", "σ", "q", "T", "mu", "n", "m", "pct"}.ToColumn();
    }
}
