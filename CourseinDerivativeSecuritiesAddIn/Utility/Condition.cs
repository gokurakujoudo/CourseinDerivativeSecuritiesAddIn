using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Utility {
    public static class Condition {
        [ExcelFunction(Category = "Conditional")]
        public static object IfEqual(double input, double condition, object ifTrue, object ifFalse) =>
            input == condition ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfLarger(double input, double condition, object ifTrue, object ifFalse) =>
            input > condition ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfLargerOrEqual(double input, double condition, object ifTrue, object ifFalse) =>
            input >= condition ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfSmaller(double input, double condition, object ifTrue, object ifFalse) =>
            input < condition ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfSmallerOrEqual(double input, double condition, object ifTrue, object ifFalse) =>
            input <= condition ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfContains(string input, string condition, object ifTrue, object ifFalse) =>
            input.Contains(condition) ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object IfContainsIgnoreCases(string input, string condition, object ifTrue, object ifFalse) =>
            input.ToLower().Contains(condition.ToLower()) ? ifTrue : ifFalse;

        [ExcelFunction(Category = "Conditional")]
        public static object SwitchBetween(double input, double[] condition, object[] output, object def = null) {
            for (var lower = condition.Length - 1; lower >= 0; lower--)
                if (input >= condition[lower])
                    return output[lower];
            return def ?? $"Not found {input}";
        }

        [ExcelFunction(Category = "Conditional")]
        public static object SwitchCases(object input, object[] condition, object[] output, object def = null)
        {
            for (var i = 0; i < condition.Length; i++)
                if (input.Equals(condition[i]))
                    return output[i];
            return def ?? $"Not found {input}";
        }
    }
}
