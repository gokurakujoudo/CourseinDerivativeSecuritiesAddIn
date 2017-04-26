using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DerivativeSecuritiesAddIn.Alpha {
    public class SharpObject {
        internal Range Init { get; }
        internal Dictionary<string, object> Properties { get; }
        internal Dictionary<string, Func<SharpObject>> ComputedProperties { get; }

        public SharpObject(Range init, Dictionary<string, object> value, Action<SharpObject> temp = null) {
            this.Init = init;
            this.Properties = new Dictionary<string, object>(value);
            this.ComputedProperties = new Dictionary<string, Func<SharpObject>>();
            temp?.Invoke(this);
        }

        public override string ToString() => $"{this.Id()} : {this.Type()} [{this.Init.ToAddress()}]";

        internal object this[string key] {
            get {
                if (this.Properties.TryGetValue(key, out var v)) return v;
                return this.ComputedProperties.TryGetValue(key, out var f) ? f() : null;
            }
            set {
                if (value is Func<SharpObject> a) {
                    this.ComputedProperties[key] = a;
                    if (this.Properties.ContainsKey(key))
                        this.Properties.Remove(key);
                }
                else {
                    this.Properties[key] = value;
                    if (this.ComputedProperties.ContainsKey(key))
                        this.ComputedProperties.Remove(key);
                }
            }
        }
    }

    public static class SharpObjectHelper {
        private static readonly Application App = (Application) ExcelDnaUtil.Application;
        private static readonly Dictionary<string, SharpObject> Env = new Dictionary<string, SharpObject>();
        private static readonly Dictionary<string, Action<SharpObject>> Templates = FindTemplates();
        private const string ID = "OBJ ID";
        private const string TYPE = "TYPE";
        private const string UNDIFINED = "UNDIFINED";
        private static readonly string[] Spefields = {ID, TYPE};
        internal const string TEMPPROVIDERMETHOD = "ProvideTemplate";


        internal static object PrintCompulted(this SharpObject obj) {
            var n = obj.ComputedProperties.Count;
            var ar = obj.ComputedProperties.ToArray();
            var result = new object[n, 2];
            for (var i = 0; i < n; i++) {
                result[i, 0] = ar[i].Key.ToUpper();
                result[i, 1] = (object) ar[i].Value?.Invoke() ?? UNDIFINED;
            }
            return result;
        }

        internal static object Print(this SharpObject obj) {
            var dict = obj.Properties;
            var id = dict[ID].To<string>();
            dict.TryGetValue(TYPE, out var type);
            type = type ?? UNDIFINED;
            var prop = obj.Properties.Where(p => !Spefields.Contains(p.Key)).ToArray();
            var n = prop.Length;
            var cprop = obj.ComputedProperties.ToArray();
            var m = cprop.Length;
            var result = new object[n + m + 2, 2];
            result[0, 0] = ID;
            result[0, 1] = id;
            result[1, 0] = TYPE;
            result[1, 1] = type;
            for (var i = 0; i < n; i++) {
                result[i + 2, 0] = prop[i].Key.ToUpper();
                var v = prop[i].Value ?? UNDIFINED;
                if (v is Range r)
                    result[i + 2, 1] = r.ToAddress();
                result[i + 2, 1] = v;
            }
            for (var i = 0; i < m; i++) {
                result[i + 2 + n, 0] = cprop[i].Key.ToUpper();
                var v = (object) cprop[i].Value?.Invoke() ?? UNDIFINED;
                if (v is Range r)
                    result[i + 2 + n, 1] = r.ToAddress();
                result[i + 2 + n, 1] = v;
            }
            return result;
        }


        [ExcelFunction(Category = "Alpha", IsVolatile = true, IsMacroType = true)]
        public static object CreateObj(object[,] value) {
            var rref = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            var range = rref.ToRange();
            var dict = value.ToDict();
            if (!dict.TryGetValue(ID, out var rid)) return "Invalid Selection";
            if (rid is Range || string.IsNullOrWhiteSpace(rid.ToString()))
                return "Invalid Id";
            var id = $"@{rid.ToString().ToUpper()}";
            if (Env.TryGetValue(id, out var same) && same.Init.ToAddress() != range.ToAddress())
                if (same.Init.Value == id)
                    return $"Duplicate Id in {same.Init.ToAddress()}";
            Action<SharpObject> temp = null;
            if (dict.TryGetValue(TYPE, out var rtemp)) {
                if (rtemp is Range || string.IsNullOrWhiteSpace(rtemp.ToString()))
                    return "Invalid Type";
                if (!Templates.TryGetValue(rtemp.ToString(), out temp))
                    return "Invalid Type";
            }
            var so = new SharpObject(range, dict, temp) {[ID] = id};
            Env[id] = so;
            return id;
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object ViewObj(string id) {
            Env.TryGetValue(id, out var k);
            return k.Print();
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object ViewObjComputedProp(string id) {
            Env.TryGetValue(id, out var k);
            return k.PrintCompulted();
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true)]
        public static object ViewEnvIds(string type = "") {
            var d = Env.Values;
            var strs = Env.Values
                          .Where(s => string.IsNullOrWhiteSpace(type) || s.Type() == type)
                          .Select(s => s.ToString())
                          .ToArray();
            return strs.ToColumn();
        }

        internal static string Id(this SharpObject so) { return so[ID].ToString(); }
        internal static string Type(this SharpObject so) { return so[TYPE]?.ToString() ?? UNDIFINED; }

        private static Dictionary<string, Action<SharpObject>> FindTemplates() {
            //TODO
            return new Dictionary<string, Action<SharpObject>>();
        }
    }
}
