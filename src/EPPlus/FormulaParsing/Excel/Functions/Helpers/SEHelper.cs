using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class SEHelper
    {
        public static double GetStandardError(double[] xValues, double[] yValues, bool pushToZero)
        {

            double yMean = yValues.Average();
            double xMean = xValues.Average();
            int sampleSize = yValues.Count();
            var p1 = 0d;
            var numerator = 0d;
            var denominator = 0d;

            for (var i = 0; i < yValues.Count(); i++)
            {
                double y1 = yValues[i];
                double x1 = xValues[i];

                if (pushToZero)
                {
                    p1 += Math.Pow(y1, 2);
                    numerator += x1 * y1;
                    denominator += Math.Pow(x1, 2);
                }
                else
                {
                    p1 += System.Math.Pow(y1 - yMean, 2);
                    numerator += (x1 - xMean) * (y1 - yMean);
                    denominator += (System.Math.Pow(x1 - xMean, 2));
                }
            }

            double result = (pushToZero) ? Math.Sqrt((p1 - numerator * numerator / denominator) / (sampleSize - 1)) :
                                           Math.Sqrt((p1 - numerator * numerator / denominator) / (sampleSize - 2));

            return result;
        }

        public static double DevSq(double[] array, bool meanIsZero)
        {
            //Returns the sum of squares of deviations from a set of datapoints.
            var mean = (!meanIsZero) ? array.Select(x => (double)x).Average() : 0d;
            return array.Aggregate(0d, (val, x) => val += Math.Pow(x - mean, 2));
        }
    }
}