using System;
using System.Collections.Generic;
using System.Linq;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.Alpha {
    public class SharpObject {
        private static readonly Application App = (Application)ExcelDnaUtil.Application;

        private static readonly Dictionary<string, object> Env = new Dictionary<string, object>();

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object CreateObj(string id, double n1, string s1) {
            var k = new object[] {n1, s1};
            Env[id] = k;
            return id;
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object ViewObj(string id) {
            Env.TryGetValue(id, out var k);
            return k;
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object ViewEnvIds() {
            var k =  Env.Keys.ToArray();
            return k.ToColumn();
        }
    }
}
