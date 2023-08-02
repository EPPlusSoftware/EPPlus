using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class SEHelper
    {
        public static double GetStandardError(List<double> xValues, List<double> yValues)
        {

            double yMean = yValues.Average();
            double xMean = xValues.Average();
            int sampleSize = yValues.Count;
            var p1 = 0d;
            var numerator = 0d;
            var denominator = 0d;

            for (var i = 0; i < yValues.Count; i++)
            {
                double y1 = yValues[i];
                double x1 = xValues[i];

                p1 += System.Math.Pow(y1 - yMean, 2);
                numerator += (x1 - xMean) * (y1 - yMean);
                denominator += (System.Math.Pow(x1 - xMean, 2));
            }

            double result = System.Math.Sqrt((p1 - numerator * numerator / denominator) / (sampleSize - 2));

            return result;
        }
    }
}
