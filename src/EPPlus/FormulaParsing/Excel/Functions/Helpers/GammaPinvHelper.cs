/*********************************************************
 
Copyright (c) 2013 jStat

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           Ported from JavaScript to C#
 *************************************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class GammaPinvHelper
    {
        public static double gammapinv(double p, double a)
        {
            var j = 0;
            var a1 = a - 1;
            var EPS = 1e-8;
            var gln = GammaHelper.logGamma(a);
            double x;
            double err;
            double t;
            double u;
            double pp;
            double lna1 = 0;
            double afac = 0;

            if (p >= 1)
                return System.Math.Max(100, a + 100 * System.Math.Sqrt(a));
            if (p <= 0)
                return 0;
            if (a > 1)
            {
                lna1 = System.Math.Log(a1);
                afac = System.Math.Exp(a1 * (lna1 - 1) - gln);
                pp = (p < 0.5) ? p : 1 - p;
                t = System.Math.Sqrt(-2 * System.Math.Log(pp));
                x = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t;
                if (p < 0.5)
                    x = -x;
                x = System.Math.Max(1e-3,
                             a * System.Math.Pow(1 - 1 / (9 * a) - x / (3 * System.Math.Sqrt(a)), 3));
            }
            else
            {
                t = 1 - a * (0.253 + a * 0.12);
                if (p < t)
                    x = System.Math.Pow(p / t, 1 / a);
                else
                    x = 1 - System.Math.Log(1 - (p - t) / (1 - t));
            }

            for (; j < 12; j++)
            {
                if (x <= 0)
                    return 0;
                err = GammaHelper.regularizedGammaP(a, x, 1.0e-15, 10000) - p;
                //err = jStat.lowRegGamma(a, x) - p;
                if (a > 1)
                    t = afac * System.Math.Exp(-(x - a1) + a1 * (System.Math.Log(x) - lna1));
                else
                    t = System.Math.Exp(-x + a1 * System.Math.Log(x) - gln);
                u = err / t;
                x -= (t = u / (1 - 0.5 * System.Math.Min(1, u * ((a - 1) / x - 1))));
                if (x <= 0)
                    x = 0.5 * (x + t);
                if (System.Math.Abs(t) < EPS * x)
                    break;
            }

            return x;
        }
    }
}
