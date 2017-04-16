using System;

namespace CourseinDerivativeSecuritiesAddIn
{
    internal static class Util {
        public static T To<T>(this object obj) => (T) obj;

        public static object[,] ToColumn(this object[] input)
        {
            var l = input.Length;
            var output = new object[l, 1];
            for (var i = 0; i < l; i++)
                output[i, 0] = input[i];
            return output;
        }

        public const double TOL = 1E-6;

        public static double Percentile(double[] sequence, double excelPercentile)
        {
            Array.Sort(sequence);
            var length = sequence.Length;
            var n = (length - 1) * excelPercentile + 1;
            if (n == 1d) return sequence[0];
            if (n == length) return sequence[length - 1];
            var k = (int)n;
            var d = n - k;
            return sequence[k - 1] + d * (sequence[k] - sequence[k - 1]);
        }
    }
}
