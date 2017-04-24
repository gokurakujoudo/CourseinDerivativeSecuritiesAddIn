using System;
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
                      let efa = method.GetCustomAttribute(typeof(ExcelFunctionAttribute)).To<ExcelFunctionAttribute>()
                      where efa != null && efa.Category.Contains(space)
                      let paras = method.GetParameters()
                                        .Select(p => p.IsOptional
                                                    ? $"[{p.ParameterType.Name} {p.Name}]"
                                                    : $"{p.ParameterType.Name} {p.Name}")
                                        .ToArray()
                      let pstr = $"{string.Join(", ", paras)}"
                      let des = string.IsNullOrWhiteSpace(efa.Description) ? "No comment" : efa.Description
                      orderby efa.Category, method.Name
                      select new[] {efa.Category, method.Name, pstr, des}).ToList();
            var result = new object[ls.Count, 4];
            for (var i = 0; i < result.GetLength(0); i++)
            for (var j = 0; j < 4; j++)
                result[i, j] = ls[i][j];
            return result;
        }

        public struct ParaType {
            public string Name;
            public Type Type;
            public bool Optional;
            public object Default;
        }

        internal static ParaType[] GetParaInfo(this MethodInfo method) {
            return method.GetParameters()
                         .Select(p => {
                             var at = p.GetCustomAttribute(typeof(ExNameAttribute)) as ExNameAttribute;
                             return new ParaType {
                                 Name = at?.Name ?? p.Name.ToUpper(),
                                 Type = p.ParameterType,
                                 Optional = p.IsOptional,
                                 Default = p.DefaultValue
                             };
                         })
                         .ToArray();
        }
    }
}
