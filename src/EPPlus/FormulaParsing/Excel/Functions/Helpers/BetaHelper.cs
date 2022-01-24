/*
 * MIT License
 * 
 * Copyright (c) [year] [fullname]
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:

 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.

 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *************************************************************************************************
 * Date               Author                       Change
 *************************************************************************************************
 * 01/02/2021         EPPlus Software AB         Ported code from JavaScript to C# (https://github.com/jstat/jstat/blob/1.x/dist/jstat.js)
 *************************************************************************************************
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class BetaHelper
    {
        internal static double IBetaInv(double p, double a, double b)
        {
            var EPS = 1e-8;
            var a1 = a - 1;
            var b1 = b - 1;
            var j = 0;
            double lna, lnb, pp, t, u, err, x, al, h, w, afac;
            if (p <= 0)
                return 0;
            if (p >= 1)
                return 1;
            if (a >= 1 && b >= 1)
            {
                pp = (p < 0.5) ? p : 1 - p;
                t = System.Math.Sqrt(-2 * System.Math.Log(pp));
                x = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t;
                if (p < 0.5)
                    x = -x;
                al = (x * x - 3) / 6;
                h = 2 / (1 / (2 * a - 1) + 1 / (2 * b - 1));
                w = (x * System.Math.Sqrt(al + h) / h) - (1 / (2 * b - 1) - 1 / (2 * a - 1)) *
                    (al + 5 / 6 - 2 / (3 * h));
                x = a / (a + b * System.Math.Exp(2 * w));
            }
            else
            {
                lna = System.Math.Log(a / (a + b));
                lnb = System.Math.Log(b / (a + b));
                t = System.Math.Exp(a * lna) / a;
                u = System.Math.Exp(b * lnb) / b;
                w = t + u;
                if (p < t / w)
                    x = System.Math.Pow(a * w * p, 1 / a);
                else
                    x = 1 - System.Math.Pow(b * w * (1 - p), 1 / b);
            }
            //afac = -jStat.gammaln(a) - jStat.gammaln(b) + jStat.gammaln(a + b);
            afac = -GammaHelper.logGamma(a) - GammaHelper.logGamma(b) + GammaHelper.logGamma(a + b);
            for (; j < 10; j++)
            {
                if (x == 0 || x == 1)
                    return x;
                err = IBeta(x, a, b) - p;
                t = System.Math.Exp(a1 * System.Math.Log(x) + b1 * System.Math.Log(1 - x) + afac);
                u = err / t;
                x -= (t = u / (1 - 0.5 * System.Math.Min(1, u * (a1 / x - b1 / (1 - x)))));
                if (x <= 0)
                    x = 0.5 * (x + t);
                if (x >= 1)
                    x = 0.5 * (x + t + 1);
                if (System.Math.Abs(t) < EPS * x && j > 0)
                    break;
            }
            return x;
        }

        /// <summary>
        /// Returns the inverse of the incomplete beta function
        /// </summary>
        /// <param name="x"></param>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        internal static double IBeta(double x, double a, double b)
        {
            // Factors in front of the continued fraction.
            var bt = (x == 0 || x == 1) ? 0 :
              System.Math.Exp(GammaHelper.logGamma(a + b) - GammaHelper.logGamma(a) -
                       GammaHelper.logGamma(b) + a * System.Math.Log(x) + b *
                       System.Math.Log(1 - x));
            if (x < 0 || x > 1)
                return 0d; // previously return false
            if (x < (a + 1) / (a + b + 2))
                // Use continued fraction directly.
                return bt * BetaCf(x, a, b) / a;
            // else use continued fraction after making the symmetry transformation.
            return 1 - bt * BetaCf(1 - x, b, a) / b;
        }

        internal static double Beta(double x, double y)
        {
            // ensure arguments are positive
            if (x <= 0 || y <= 0)
                return 0;
            // make sure x + y doesn't exceed the upper limit of usable values
            return (x + y > 170)
                ? System.Math.Exp(Betaln(x, y))
                : GammaHelper.gamma(x) * GammaHelper.gamma(y) / GammaHelper.gamma(x + y);
        }

        internal static double Betaln(double x, double y)
        {
            return GammaHelper.logGamma(x) + GammaHelper.logGamma(y) - GammaHelper.logGamma(x + y);
        }

        internal static double BetaCdf(double x, double a, double b)
        {
            if( x > 1 || x < 0)
            {
                return x > 1 ? 1 : 0;
            }
            return IBeta(x, a, b);
        }

        internal static double BetaPdf(double x, double a, double b)
        {
            if (x > 1 || x < 0)
                return 0;
            // PDF is one for the uniform case
            if (a == 1 && b == 1)
                return 1;

            if (a < 512 && b < 512)
            {
                var result = (System.Math.Pow(x, a - 1) * System.Math.Pow(1 - x, b - 1)) /
                    Beta(a, b);
                return result / 2d;
            }
            else
            {
                var result = System.Math.Exp((a - 1) * System.Math.Log(x) +
                                (b - 1) * System.Math.Log(1 - x) -
                                Betaln(a, b));
                return result / 2d;
            }
        }

        /// <summary>
        /// Evaluates the continued fraction for incomplete beta function by modified Lentz's method.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        internal static double BetaCf(double x, double a, double b)
        {
            var fpmin = 1e-30;
            var m = 1;
            var qab = a + b;
            var qap = a + 1;
            var qam = a - 1;
            var c = 1d;
            double d = 1 - qab * x / qap;
            double m2, aa, del, h;

            // These q's will be used in factors that occur in the coefficients
            if (System.Math.Abs(d) < fpmin)
                d = fpmin;
            d = 1 / d;
            h = d;

            for (; m <= 100; m++)
            {
                m2 = 2 * m;
                aa = m * (b - m) * x / ((qam + m2) * (a + m2));
                // One step (the even one) of the recurrence
                d = 1 + aa * d;
                if (System.Math.Abs(d) < fpmin)
                    d = fpmin;
                c = 1d + aa / c;
                if (System.Math.Abs(c) < fpmin)
                    c = fpmin;
                d = 1 / d;
                h *= d * c;
                aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
                // Next step of the recurrence (the odd one)
                d = 1 + aa * d;
                if (System.Math.Abs(d) < fpmin)
                    d = fpmin;
                c = 1 + aa / c;
                if (System.Math.Abs(c) < fpmin)
                    c = fpmin;
                d = 1 / d;
                del = d * c;
                h *= del;
                if (System.Math.Abs(del - 1.0) < 3e-7)
                    break;
            }

            return h;
        }
    }
}
