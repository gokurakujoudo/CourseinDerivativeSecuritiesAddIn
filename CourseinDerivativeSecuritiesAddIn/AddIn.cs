using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace DerivativeSecuritiesAddIn {
    public class AddIn : IExcelAddIn {
        public void AutoOpen() {
            IntelliSenseServer.Register();
        }

        public void AutoClose() {

        }
    }
}