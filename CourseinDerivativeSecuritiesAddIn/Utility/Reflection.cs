using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Utility {

    public static class Reflection {
        [ExcelFunction(Category = "Utility")]
        public static object GetFunctionList(string space = "") {
            var ls = (from type in Assembly.GetExecutingAssembly().GetTypes()
                      where type.Namespace != null && type.Namespace.Contains("DerivativeSecuritiesAddIn")
                      from method in type.GetMethods()
                      let efa = method.GetCustomAttribute<ExcelFunctionAttribute>()
                      where efa != null && efa.Category.Contains(space)
                      let paras = method.GetParameters()
                                        .Select(p => p.IsOptional
                                                    ? $"[{p.ParameterType.Name} {p.Name}]"
                                                    : $"{p.ParameterType.Name} {p.Name}")
                                        .ToArray()
                      let pstr = $"{string.Join(", ", paras)}"
                      let des = string.IsNullOrWhiteSpace(efa.Description) ? "No comment" : efa.Description
                      orderby efa.Category, method.Name
                      select new[] {efa.Category, type.Name, method.Name, pstr, des}).ToList();
            var result = new object[ls.Count, 5];
            for (var i = 0; i < result.GetLength(0); i++)
            for (var j = 0; j < 5; j++)
                result[i, j] = ls[i][j];
            return result;
        }

        internal struct ParaType {
            public string Name;
            public Type Type;
            public bool Optional;
            public object Default;
        }

        internal static ParaType[] GetParaInfo(this MethodInfo method) {
            return method.GetParameters()
                         .Select(p => {
                             var at = p.GetCustomAttribute<ExNameAttribute>();
                             return new ParaType {
                                 Name = at?.Name ?? p.Name.ToUpper(),
                                 Type = p.ParameterType,
                                 Optional = p.IsOptional,
                                 Default = p.DefaultValue
                             };
                         })
                         .ToArray();
        }

        private static readonly string[] TempSpace = {
            "Binomial", "BlackScholes", "Corporate Finance",
            "Exotic Option", "Generic Option", "Simulation"
        };

        internal static Dictionary<string, List<string>> GetTempsStrs() => (
            from type in Assembly.GetExecutingAssembly().GetTypes()
            where type.Namespace != null && type.Namespace.Contains("DerivativeSecuritiesAddIn")
            from method in type.GetMethods()
            let efa = method.GetCustomAttribute<ExcelFunctionAttribute>()
            where efa != null && TempSpace.Contains(efa.Category)
            let print = CutName(method.Name)
            orderby efa.Category, print
            group print by type.Name into g
            select g).ToDictionaryEx(g => g.Key, g => g.ToList());

        internal static Dictionary<string, Action> GetTempsActions() => (
            from type in Assembly.GetExecutingAssembly().GetTypes()
            where type.Namespace != null && type.Namespace.Contains("DerivativeSecuritiesAddIn")
            from method in type.GetMethods()
            let efa = method.GetCustomAttribute<ExcelFunctionAttribute>()
            where efa != null && TempSpace.Contains(efa.Category)
            select method).ToDictionaryEx<MethodInfo, string, Action>(
            m => CutName(m.Name).ToUpper(), m => () => TemplateSubBase.CreateTemp(m));

        internal static Dictionary<string, MethodInfo> GetTempsMethods() => (
            from type in Assembly.GetExecutingAssembly().GetTypes()
            where type.Namespace != null && type.Namespace.Contains("DerivativeSecuritiesAddIn")
            from method in type.GetMethods()
            let efa = method.GetCustomAttribute<ExcelFunctionAttribute>()
            where efa != null 
            select method).ToDictionaryEx(m => CutName(m.Name).ToUpper(), m => m);

        internal static string CutName(string methodName) {
            if (!methodName.Contains("_")) return methodName;
            for (var i = 0; i < methodName.Length; i++) 
                if (methodName[i] == '_')
                    return methodName.Substring(0, i);
            return methodName;
        }
    }
}
