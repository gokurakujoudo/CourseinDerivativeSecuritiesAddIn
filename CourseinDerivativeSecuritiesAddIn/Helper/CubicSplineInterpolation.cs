using System;
using System.Linq;
using DerivativeSecuritiesAddIn.Util;
using ExcelDna.Integration;

namespace DerivativeSecuritiesAddIn.Helper
{
    internal enum EnmBoradType {
        SecondOrder = 0,
        FirstOrder = 1,
        [Obsolete] Periodic = 2
    }

    public static class CubicSplineInterpolation {

        [ExcelFunction(Category = "Interpolation")]
        public static object[,] InCubicSplineFit2Ord(double[] xs, double[] ys, double[] newx, double d, double u) {
            var cur = CubicInt(xs, ys, EnmBoradType.SecondOrder, d, u);
            var newy = newx.Select(x => cur.Fit(x)).ToArray();
            return newy.ToColumn();
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object[,] InCubicSplineFit1Ord(double[] xs, double[] ys, double[] newx, double d, double u)
        {
            var cur = CubicInt(xs, ys, EnmBoradType.FirstOrder, d, u);
            var newy = newx.Select(x => cur.Fit(x)).ToArray();
            return newy.ToColumn();
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object InCubicFirstOrder(double[] xs, double[] ys) {
            var n = xs.Length;
            var dx = (xs[n - 1] - xs[0]) / 10000;
            var cur = CubicInt(xs, ys, EnmBoradType.SecondOrder, 0, 0);

            var delta = new double[n];
            var y0Po = cur.Fit(xs[0] + dx);
            delta[0] = (y0Po - ys[0]) / dx;
            var ynNe = cur.Fit(xs[n - 1] - dx);
            delta[n - 1] = (ys[n - 1] - ynNe) / dx;
            for (var i = 1; i < n - 1; i++) {
                var po = cur.Fit(xs[i] + dx);
                var ne = cur.Fit(xs[i] - dx);
                delta[i] = 0.5 * (po - ne) / dx;
            }
            return delta.ToColumn();
        }

        [ExcelFunction(Category = "Interpolation")]
        public static object InCubicExpand_2xn(double[] xs, double[] ys) {
            var range = XlCall.Excel(XlCall.xlfCaller).To<ExcelReference>();
            var count = range.RowLast - range.RowFirst + 1;
            if (count == 1)
                return "This is an array function";
            var newx = LinearInterpolation.LinearSeq_internal(xs[0], xs[xs.Length - 1], count);
            var cur = CubicInt(xs, ys, EnmBoradType.SecondOrder, 0, 0);
            var newy = newx.Select(x => cur.Fit(x)).ToArray();
            var result = new object[count, 2];
            for (var i = 0; i < count; i++) {
                result[i, 0] = newx[i];
                result[i, 1] = newy[i];
            }
            return result;
        }

        internal static double Fit(this object[][] cubic, double x0) =>
            cubic.First(tu => x0 >= tu[0].To<double>() && x0 <= tu[1].To<double>())[2]
                 ?.To<Func<double, double>>()
                 .Invoke(x0) ?? throw new ArgumentOutOfRangeException();

        //http://blog.csdn.net/zhangxiaolu2015/article/details/42744823
        internal static object[][] CubicInt(double[] x, double[] y, EnmBoradType type, params double[] args) {
            var n = x.Length;
            //n=size(x,2);  

            //for k=2:n 
            //    h(k)=x(k)-x(k-1);
            //end
            var h = new double[n];
            for (var k = 1; k < n; k++)
                h[k] = x[k] - x[k - 1];

            //for k=2:(n-1) 
            //   mu(k)=h(k)/(h(k)+h(k+1));
            //   lambda(k)=1-mu(k);
            //end
            var mu = new double[n];
            var lambda = new double[n];
            for (var k = 1; k < n - 1; k++) {
                mu[k] = h[k] / (h[k] + h[k + 1]);
                lambda[k] = 1 - mu[k];
            }

            //for k=2:(n-1)
            //    d(k)=6*((y(k+1)-y(k))/h(k+1)-(y(k)-y(k-1))/h(k))/(h(k)+h(k+1)); 
            //end
            var d = new double[n];
            for (var k = 1; k < n - 1; k++)
                d[k] = 6 * ((y[k + 1] - y[k]) / h[k + 1] - (y[k] - y[k - 1]) / h[k]) / (h[k] + h[k + 1]);

            var m = new double[n];
            switch (type) {
                case EnmBoradType.SecondOrder:
                    m[0] = args[0];
                    m[n - 1] = args[1];

                    //    A=zeros(n-2,n-2);
                    var a = new double[n - 2, n - 2];

                    //    for k=1:(n-3) 
                    //        A(k,k)=2;
                    //        A(k,k+1)=lambda(k+1);
                    //        A(k+1,k)=mu(k+2);
                    //    end
                    for (var k = 0; k < n - 3; k++) {
                        a[k, k] = 2;
                        a[k, k + 1] = lambda[k + 1];
                        a[k + 1, k] = mu[k + 2];
                    }

                    //    A(n-2,n-2)=2;
                    a[n - 3, n - 3] = 2;

                    //    b=zeros(n-2,1);
                    var b = new double[n - 2];

                    //    for k=2:(n-3)
                    //    b(k,1)=d(k+1);
                    //    end
                    for (var k = 1; k < n - 3; k++)
                        b[k] = d[k + 1];

                    //    b(1,1)=d(2)-mu(2)*M(1);
                    //    b(n-2,1)=d(n-1)-lambda(n-1)*M(n);
                    b[0] = d[1] - mu[1] * m[0];
                    b[n - 3] = d[n - 2] - lambda[n - 2] * m[n - 1];

                    //    N=ZhuiGanFa(A,b);            
                    var solution = LinearAlgebraHelper.CatchSolver(a, b);

                    //    for k=2:(n-1)
                    //        M(k)=N(k-1,1);
                    //    end
                    for (var k = 1; k < n - 1; k++)
                        m[k] = solution[k - 1];

                    break;
                case EnmBoradType.FirstOrder:
                    var y0 = args[0];
                    var yn = args[0];

                    //    d(1)=6*((y(2)-y(1))/h(2)-y0)/h(2);
                    //    d(n)=6*(yn-(y(n)-y(n-1))/h(n))/h(n);
                    d[0] = 6 * ((y[1] - y[0]) / h[1] - y0) / h[1];
                    d[n - 1] = 6 * (yn - (y[n - 1] - y[n - 2]) / h[n - 1]) / h[n - 1];

                    //    A=zeros(n,n);
                    var a2 = new double[n, n];

                    //    for k=2:(n-2) 
                    //            A(k,k)=2;
                    //            A(k,k+1)=lambda(k);
                    //            A(k+1,k)=mu(k+1);
                    //    end
                    for (var k = 1; k < n - 2; k++) {
                        a2[k, k] = 2;
                        a2[k, k + 1] = lambda[k];
                        a2[k + 1, k] = mu[k + 1];
                    }

                    a2[0, 0] = 2; // A(1,1)=2;
                    a2[0, 1] = 1; // A(1,2)=1;
                    a2[1, 0] = mu[1]; // A(2,1)=mu(2);
                    a2[n - 2, n - 2] = 2; // A(n-1,n-1)=2;
                    a2[n - 1, n - 1] = 2; // A(n,n)=2;
                    a2[n - 1, n - 2] = 1; // A(n,n-1)=1;
                    a2[n - 2, n - 1] = lambda[n - 2]; // A(n-1,n)=lambda(n-1);

                    //        b=zeros(n,1);
                    var b2 = new double[n];

                    //        for k=1:n
                    //            b(k,1)=d(k);
                    //        end
                    for (var k = 0; k < n; k++)
                        b2[k] = d[k];

                    //        N=ZhuiGanFa(A,b);
                    var solution2 = LinearAlgebraHelper.CatchSolver(a2, b2);

                    //        for k=1:n
                    //            M(k)=N(k,1);
                    //        end
                    for (var k = 0; k < n; k++)
                        m[k] = solution2[k];

                    break;
/*                case EnmBoradType.Periodic:
                    //        d(n)=6*((y(2)-y(n))/h(2)-(y(n)-y(n-1))/h(n))/(h(n)+h(2));
                    d[n - 1] = 6 * ((y[1] - y[n - 1]) / h[1] - (y[n - 1] - y[n - 2]) / h[n - 1]) / (h[n - 1] + h[1]);

                    //        A=zeros(n-1,n-1);
                    var A3 = new double[n - 1, n - 1];

                    //        for k=1:(n-2)
                    //            A(k,k)=2;
                    //            A(k,k+1)=lambda(k+1);
                    //            A(k+1,k)=mu(k+2);
                    //        end
                    for (var k = 0; k < n - 2; k++) {
                        A3[k, k] = 2;
                        A3[k, k + 1] = lambda[k + 1];
                        A3[k + 1, k] = mu[k + 2];
                    }

                    A3[n - 2, n - 2] = 2; // A(n-1,n-1)=2;
                    A3[0, n - 2] = mu[1]; // A(1,n-1)=mu(2);
                    A3[n - 2, 0] = lambda[n - 1]; // A(n-1,1)=lambda(n);

                    //        b=zeros(n-1,1);
                    var b3 = new double[n - 1];

                    //        for k=1:(n-1)
                    //            b(k,1)=d(k+1);
                    //        end
                    for (var k = 0; k < n - 1; k++)
                        b3[k] = d[k + 1];

                    //        N=LU_fenjieqiuxianxingfangcheng(A,b); 
                    var N3 = LinearAlgberaHelper.LuSolver(A3, b3);

                    //        for k=1:(n-1)
                    //            M(k+1)=N(k,1);
                    //        end
                    for (var k = 0; k < n - 1; k++)
                        m[k + 1] = N3[k];

                    //        M(1)=M(n);
                    m[0] = m[n - 1];

                    break;*/
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }
            //M;

            var result = new object[n - 1][];
            //for k=1:(n-1)
            for (var k = 0; k < n - 1; k++) {
                //    clear S1
                //    syms X
                //    S1=(x(k+1)-X)^3*M(k)/(6*h(k+1))+
                //        (X-x(k))^3*M(k+1)/(6*h(k+1))+...
                //        (y(k)-h(k+1)^2*M(k)/6)*(x(k+1)-X)/h(k+1)+...
                //        (y(k+1)-h(k+1)^2*M(k+1)/6)*(X-x(k))/h(k+1);
                var k1 = k;
                Func<double, double> s1 = x0 => (x[k1 + 1] - x0).Pow(3) * m[k1] / (6 * h[k1 + 1]) +
                                               (x0 - x[k1]).Pow(3) * m[k1 + 1] / (6 * h[k1 + 1]) +
                                               (y[k1] - h[k1 + 1].Pow() * m[k1] / 6) * (x[k1 + 1] - x0) / h[k1 + 1] +
                                               (y[k1 + 1] - h[k1 + 1].Pow() * m[k1 + 1] / 6) * (x0 - x[k1]) / h[k1 + 1];
                //    fprintf('当%d=<X=<%d时\n',x(k),x(k+1));
                //    S=expand(S1)
                result[k] = new object[] {x[k], x[k + 1], s1};
                //end
            }

            return result;
        }
    }
}
