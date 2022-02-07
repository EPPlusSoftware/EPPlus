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
    internal class ContinuedFraction
    {
        private const double DEFAULT_EPSILON = 10e-9;

        public Func<int, double, double> GetA
        {
            get; set;
        }

        public Func<int, double, double> GetB
        {
            get; set;
        }

        private static bool PrecisionEquals(double x, double y, double eps)
        {
            double diff = System.Math.Abs(x * eps);
            return System.Math.Abs(x - y) <= diff;
        }

        /// <summary>
        /// Evaluates the continued fraction at the value x
        /// </summary>
        /// <param name="x"></param>
        /// <returns></returns>
        public double Evaluate(double x)
        {
            return Evaluate(x, DEFAULT_EPSILON, int.MaxValue);
        }

        public double Evaluate(double x, int maxIterations)
        {
            return Evaluate(x, DEFAULT_EPSILON, maxIterations);
        }

        public double Evaluate(double x, double epsilon, int maxIterations)
        {
            const double small = 1e-50;
            double hPrev = GetA.Invoke(0, x);

            // use the value of small as epsilon criteria for zero checks
            if (PrecisionEquals(hPrev, 0.0, small))
            {
                hPrev = small;
            }

            int n = 1;
            double dPrev = 0.0;
            double cPrev = hPrev;
            double hN = hPrev;

            while (n < maxIterations)
            {
                double a = GetA.Invoke(n, x);
                double b = GetB.Invoke(n, x);

                double dN = a + b * dPrev;
                if (PrecisionEquals(dN, 0.0, small))
                {
                    dN = small;
                }
                double cN = a + b / cPrev;
                if (PrecisionEquals(cN, 0.0, small))
                {
                    cN = small;
                }

                dN = 1 / dN;
                double deltaN = cN * dN;
                hN = hPrev * deltaN;

                if (double.IsInfinity(hN))
                {
                    throw new Exception("Continued fraction infinity divergence: " + x);
                }
                if (double.IsNaN(hN))
                {
                    throw new Exception("Continued fraction NAN divergence: " + x);
                }

                if (System.Math.Abs(deltaN - 1.0) < epsilon)
                {
                    break;
                }

                dPrev = dN;
                cPrev = cN;
                hPrev = hN;
                n++;
            }

            if (n >= maxIterations)
            {
                throw new Exception("Non convergent continued fraction: " + x);
            }

            return hN;
        }
    }
}
