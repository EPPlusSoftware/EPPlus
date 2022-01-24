/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *************************************************************************************************
 * Date               Author                       Change
 *************************************************************************************************
 * 05/20/2020         EPPlus Software AB         Ported code from java to C#
 *************************************************************************************************
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class ErfHelper
    {
        private const double X_CRIT = 0.4769362762044697;

        public static double Erf(double x)
        {
            if (System.Math.Abs(x) > 40)
            {
                return x > 0 ? 1 : -1;
            }
            double ret = GammaHelper.regularizedGammaP(0.5, x * x, 1.0e-15, 10000);
            return x < 0 ? -ret : ret;
        }


        public static double Erf(double x1, double x2)
        {
            if (x1 > x2)
            {
                return -Erf(x2, x1);
            }

            return
            x1 < -X_CRIT ?
                x2 < 0.0 ?
                    Erfc(-x2) - Erfc(-x1) :
                    Erf(x2) - Erf(x1) :
                x2 > X_CRIT && x1 > 0.0 ?
                    Erfc(x1) - Erfc(x2) :
                    Erf(x2) - Erf(x1);
        }

        public static double Erfc(double x)
        {
            if (System.Math.Abs(x) > 40)
            {
                return x > 0 ? 0 : 2;
            }
            double ret = GammaHelper.regularizedGammaQ(0.5, x * x, 1.0e-15, 10000);
            return x < 0 ? 2 - ret : ret;
        }

        public static double Erfcinv(double p)
        {
            if (p >= 2)
                return -100;
            if (p <= 0)
                return 100;
            var pp = (p < 1) ? p : 2 - p;
            var t = System.Math.Sqrt(-2 * System.Math.Log(pp / 2));
            var x = -0.70711 * ((2.30753 + t * 0.27061) /
                            (1 + t * (0.99229 + t * 0.04481)) - t);
            double err;
            for (var j = 0; j < 2; j++)
            {
                err = Erfc(x) - pp;
                x += err / (1.12837916709551257 * System.Math.Exp(-x * x) - x * err);
            }
            return (p < 1) ? x : -x;
        }
    }
}
