using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Utility {
    public static class RegExp {
        [ExcelFunction(Category = "Regular Expression")]
        public static object ReMatch(string expression, string input, int index = 1) {
            var reg = new Regex(expression);
            if (!reg.IsMatch(input)) return false;
            if (index <= 1) index = 1;
            var match = reg.Matches(input)[index - 1].Value;
            return match;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ReMatchCount(string expression, string input) {
            var reg = new Regex(expression);
            if (!reg.IsMatch(input)) return false;
            var matches = reg.Matches(input);
            return matches.Count;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ReMatches(string expression, string input) {
            var reg = new Regex(expression);
            if (!reg.IsMatch(input)) return false;
            var matches = reg.Matches(input);
            var strs = new object[matches.Count];
            for (var i = 0; i < matches.Count; i++) {
                strs[i] = matches[i].Value;
            }
            return strs.ToColumn();
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ReReplace(string expression, string input, string replacement) {
            var reg = new Regex(expression);
            return reg.Replace(input, replacement);
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ExtractNumbers(string input, int index = 1) {
            var reg = new Regex(@"[0-9]+[.0-9]*");
            if (!reg.IsMatch(input)) return false;
            var matches = reg.Matches(input);
            if (index < 1) index = 1;
            return matches[index - 1].Value;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ReCountIf(string expression, object[] input) {
            var reg = new Regex(expression);
            var strs = input.Select(o => o is ExcelEmpty ? "" : o.ToString()).ToArray();
            var count = strs.Count(str => reg.IsMatch(str));
            return count;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object ReSumIf(string expression, object[] input, double[] num)
        {
            var reg = new Regex(expression);
            var strs = input.Select(o => o is ExcelEmpty ? "" : o.ToString()).ToArray();
            var sum = 0D;
            for (var i = 0; i < strs.Length; i++) 
                if (reg.IsMatch(strs[i]))
                    sum += num[i];
            return sum;
        }
    }
}
