using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DerivativeSecuritiesAddIn.ObjectSystem;
using DerivativeSecuritiesAddIn.Utility;
using Microsoft.Office.Interop.Excel;

namespace DerivativeSecuritiesAddIn.Alpha
{
    internal static class DataFrameHelper {
        internal static SharpObjectTemplate ProvideTemplate() {
            var df = new SharpObjectTemplate("DATAFRAME") {
                PropertyList = {{"COLUMNS", null}, {"DATA", null}},
                PropertyDescription = {{"COLUMNS", "Range"}, {"DATA", "Range"}},
                ComputedPropertyList = {
                    {"DATAVALUE", so => ((Range) so["DATA"]).Value},
                    {"COLUMNSVALUE", so => ((Range) so["COLUMNS"]).Value},
                    {"COLUMNSCOUNT", so => ((object[,]) so["COLUMNSVALUE"]).GetLength(1)},
                    {"ROWSCOUNT", so => ((object[,]) so["DATAVALUE"]).GetLength(0)}
                }
            };
            return df;
        }
    }
}
