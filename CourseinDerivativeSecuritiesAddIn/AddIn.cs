using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace CourseinDerivativeSecuritiesAddIn {
    public class AddIn : IExcelAddIn {
        public void AutoOpen() {
            IntelliSenseServer.Register();
        }

        public void AutoClose() {
        }
    }
}