using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DerivativeSecuritiesAddIn.Utility;

namespace DerivativeSecuritiesAddIn.Alpha
{
    internal static class DemoObjTemp {
        internal static Dictionary<string, Action<SharpObject>> ProvideTemplate() {
            var dict = new Dictionary<string, Action<SharpObject>> {
                {
                    "+1s", so => {
                        so["+1"] = so["a"].To<double>() + 1;
                    }
                }

            };
            return dict;
        }
    }
}
