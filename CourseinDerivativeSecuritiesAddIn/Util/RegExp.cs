using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Util
{
    public static class RegExp
    {
        [ExcelFunction(Category = "Regular Expression")]
        public static object RegexMatch(string expression, string input) {
            var reg = new Regex(expression);
            if (!reg.IsMatch(input)) return false;
            var match = reg.Match(input).Value;
            return match;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object RegexMatchCount(string expression, string input)
        {
            var reg = new Regex(expression);
            if (!reg.IsMatch(input)) return false;
            var matches = reg.Matches(input);
            return matches.Count;
        }

        [ExcelFunction(Category = "Regular Expression")]
        public static object RegexMatches(string expression, string input)
        {
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
        public static object RegexReplace(string expression, string input, string replacement)
        {
            var reg = new Regex(expression);
            return reg.Replace(input, replacement);
        }
    }
}
