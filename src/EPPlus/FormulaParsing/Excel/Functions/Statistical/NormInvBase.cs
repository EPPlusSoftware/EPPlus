
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/29/2021         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal abstract class NormInvBase : ExcelFunction
    {
        private const double S1 = 0.425E0;
        private const double S2 = 5;
        private const double C1 = 0.180625E0;
        private const double C2 = 1.6;

        private static readonly double[] A = new double[]
        {
            3.3871328727963666080E0,
            1.3314166789178437745E2,
            1.9715909503065514427E3,
            1.3731693765509461125E4,
            4.5921953931549871457E4,
            6.7265770927008700853E4,
            3.3430575583588128105E4,
            2.5090809287301226727E3
        };

        private static readonly double[] B = new double[]
        {
            0d,
            4.2313330701600911252E1,
            6.8718700749205790830E2,
            5.3941960214247511077E3,
            2.1213794301586595867E4,
            3.9307895800092710610E4,
            2.8729085735721942674E4,
            5.2264952788528545610E3
        };

        private static readonly double[] C = new double[]
        {
            1.42343711074968357734E0,
            4.63033784615654529590E0,
            5.76949722146069140550E0,
            3.64784832476320460504E0,
            1.27045825245236838258E0,
            2.41780725177450611770E-1,
            2.27238449892691845833E-2,
            7.74545014278341407640E-4,
        };

        private static readonly double[] D = new double[]
        {
            0d,
            2.05319162663775882187E0,
            1.67638483018380384940E0,
            6.89767334985100004550E-1,
            1.48103976427480074590E-1,
            1.51986665636164571966E-2,
            5.47593808499534494600E-4,
            1.05075007164441684324E-9
        };

        private static readonly double[] E = new double[]
        {
             6.65790464350110377720E0,
             5.46378491116411436990E0,
             1.78482653991729133580E0,
             2.96560571828504891230E-1,
             2.65321895265761230930E-2,
             1.24266094738807843860E-3,
             2.71155556874348757815E-5,
             2.01033439929228813265E-7,
        };

        private static readonly double[] F = new double[]
        {
            0d,
            5.99832206555887937690E-1,
            1.36929880922735805310E-1,
            1.48753612908506148525E-2,
            7.86869131145613259100E-4,
            1.84631831751005468180E-5,
            1.42151175831644588870E-7,
            2.04426310338993978564E-15
        };

        protected static double NormsInv(double p, double mean, double stdev)
        {

            double q, r, val;

            q = p - 0.5;

            if (System.Math.Abs(q) <= S1)
            {
                r = C1 - System.Math.Pow(q, 2);
                val = q * (((((((A[7] * r + A[6]) * r + A[5]) * r + A[4]) * r + A[3]) * r + A[2]) * r + A[1]) * r + A[0])
                       / (((((((B[7] * r + B[6]) * r + B[5]) * r + B[4]) * r + B[3]) * r + B[2]) * r + B[1]) * r + 1);
            }
            else
            {
                if (q < 0)
                    r = p;
                else
                    r = 1 - p;


                r = System.Math.Sqrt(-System.Math.Log(r));

                if (r <= S2)
                {
                    r += -C2;
                    val = (((((((C[7] * r + C[6]) * r + C[5]) * r + C[4]) * r + C[3]) * r + C[2]) * r + C[1]) * r + C[0])
                        / (((((((D[7] * r + D[6]) * r + D[5]) * r + D[4]) * r + D[3]) * r + D[2]) * r + D[1]) * r + 1);
                }
                else
                {
                    r += -5;
                    val = (((((((E[7] * r + E[6]) * r + E[5]) * r + E[4]) * r + E[3]) * r + E[2]) * r + E[1]) * r + E[0])
                        / (((((((F[7] * r + F[6]) * r + F[5]) * r + F[4]) * r + F[3]) * r + F[2]) * r + F[1]) * r + 1);
                }

                if (q < 0.0)
                {
                    val = -val;
                }
            }

            return mean + stdev * val;
        }
    }
}
