using System;
using System.Collections.Generic;
using System.Linq;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace DerivativeSecuritiesAddIn.ObjectSystem {
    internal class SharpObject {
        internal Range Init { get; }
        internal Dictionary<string, object> Properties { get; }
        internal Dictionary<string, Func<SharpObject, object>> ComputedProperties { get; }

        public SharpObject(Range init, Dictionary<string, object> value) {
            this.Init = init;
            this.Properties = new Dictionary<string, object>(value);
            this.ComputedProperties = new Dictionary<string, Func<SharpObject, object>>();
        }

        public override string ToString() => $"{this.Id()} : {this.TemplateName()} [{this.Init.ToAddress()}]";
        internal string[] ToSign() => new[] {this.Id(), this.TemplateName(), this.Init.ToAddress()};

        internal object this[string key] {
            get {
                key = key.ToUpper();
                if (this.Properties.TryGetValue(key, out var v)) return v;
                return this.ComputedProperties.TryGetValue(key, out var f) ? f(this) : null;
            }
            set {
                key = key.ToUpper();
                if (value is Func<SharpObject, object> a) {
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
        private static readonly Dictionary<string, SharpObjectTemplate> Templates = FindTemplates();
        internal const string ID = "OBJECT ID";
        internal const string TYPE = "TYPE";
        private const string UNDIFINED = "UNDIFINED";
        internal static readonly string[] Spefields = {ID, TYPE};


        internal static object[,] PrintCompulted(this SharpObject obj) {
            if (obj == null) return new object[,] {{string.Empty, "Object not found"}};
            var n = obj.ComputedProperties.Count;
            if (n == 0) return new object[,] {{string.Empty, "No Calc-Prop"}};
            var ar = obj.ComputedProperties.ToArray();
            var result = new object[n, 2];
            for (var i = 0; i < n; i++) {
                var key = ar[i].Key;
                result[i, 0] = key.ToUpper();
                result[i, 1] = obj.GetCalculatedProperty(key).ToExcelPrint();
            }
            return result;
        }

        internal static object[,] Print(this SharpObject obj) {
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
                result[i + 2, 1] = v.ToExcelPrint();
            }
            for (var i = 0; i < m; i++) {
                var key = cprop[i].Key;
                result[i + 2 + n, 0] = key.ToUpper();
                result[i + 2 + n, 1] = obj.GetCalculatedProperty(key).ToExcelPrint();
            }
            return result;
        }

        internal static SharpObjectTemplate GetTemplate(this SharpObject obj) {
            var name = obj.TemplateName();
            return Templates.TryGetValue(name, out var temp) ? temp : null;
        }

        internal static object GetCalculatedProperty(this SharpObject obj, string propertyName) {
            object value;
            try {
                value = obj.ComputedProperties[propertyName]?.Invoke(obj) ?? UNDIFINED;
            }
            catch {
                value = "ERROR";
            }
            return value;
        }

        [ExcelFunction(Category = "Alpha", IsVolatile = true, IsMacroType = true)]
        public static object CreateObjFromRef(string reference) {
            var range = reference.ToRange();
            object[,] value = range.Value;

            return CreateObj(value);
        }


        [ExcelFunction(Category = "Alpha", IsVolatile = true, IsMacroType = true)]
        public static object CreateObj(object[,] value) {
            var rref = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            var range = rref.ToRange();
            var dict = value.ToDict();
            if (!dict.TryGetValue(ID, out var rid)) return "Invalid Selection";
            if (rid ==null || rid is Range || string.IsNullOrWhiteSpace(rid.ToString()))
                return "Invalid Id";
            var id = $"@{rid.ToString().ToUpper()}";
            if (Env.TryGetValue(id, out var same) && same.Init.ToAddress() != range.ToAddress())
                if (same.Init.Value == id)
                    return $"Duplicate Id in {same.Init.ToAddress()}";
            var so = new SharpObject(range, dict) {[ID] = id};
            if (dict.TryGetValue(TYPE, out var rtemp) && !string.IsNullOrWhiteSpace(rtemp?.ToString())) {
                if (rtemp is Range)
                    return "Invalid Type";
                if (!Templates.TryGetValue(rtemp.ToString().ToUpper(), out var temp) || temp == null)
                    return "Invalid Type";
                if (temp.CheckValid(so))
                    temp.Adjust(so);
                else
                    return $"Invalid New {temp.TypeName}";
            }
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
        public static object ViewEnv(string type = "") {
            var strs = Env.Values
                          .Where(s => string.IsNullOrWhiteSpace(type) || s.TemplateName() == type)
                          .Select(s => s.ToSign())
                          .ToArray();
            var n = strs.Length;
            var results = new object[n, 3];
            for (var i = 0; i < n; i++)
            for (var j = 0; j < 3; j++)
                results[i, j] = strs[i][j];
            return results;
        }

        internal static string Id(this SharpObject so) { return so[ID].ToString(); }
        internal static string TemplateName(this SharpObject so) { return so[TYPE]?.ToString() ?? UNDIFINED; }

        private static Dictionary<string, SharpObjectTemplate> FindTemplates() {
            var dict = new Dictionary<string, SharpObjectTemplate> {
                [UNDIFINED] = new SharpObjectTemplate(UNDIFINED),
                ["DEMO"] = new SharpObjectTemplate("DEMO") {
                    PropertyList = {{"AGE", 100}},
                    PropertyDescription = {{"AGE", "double"}},
                    ComputedPropertyList = {{"NEWAGE", so => so["AGE"].ToDouble() + 1}}
                }
            };

            return dict;
        }
    }
}
