using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Template
{
    public static class GreeksAndSigns
    {
        [ExcelFunction(Category = "Greeks")]
        public static object GreeksU() =>
             "αβγδεζηθκλμνξπρστυφψω";

        [ExcelFunction(Category = "Greeks")]
        public static object GreeksC() =>
            "ΓΔΘΞΠΣΦΨΩ";
    }
}
