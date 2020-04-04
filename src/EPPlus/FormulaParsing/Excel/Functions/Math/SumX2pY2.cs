using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal class SumX2pY2 : SumxBase
    {
        public override double Calculate(double[] set1, double[] set2)
        {
            var result = 0d;
            for (var x = 0; x < set1.Length; x++)
            {
                var a = set1[x];
                var b = set2[x];
                result += a * a + b * b;
            }
            return result;
        }
    }
}
