using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class TrendHelper
    {
        internal static InMemoryRange GetTrendValuesMultiple(double[][] xRanges, double[] coefficients, bool constVar, bool columnArray)
        {
            var resultRange = (constVar) ? new InMemoryRange((short)xRanges.Count(), 1) : new InMemoryRange(1, (short)xRanges.Count());
            //If const, we remove the column of ones.
            List<double> removeColumn = [xRanges[0].Count() - 1];
            if (constVar) xRanges = MatrixHelper.RemoveColumns(xRanges, removeColumn);
            for (var i = 0; i < xRanges.Length; i++)
            {
                var trendVal = 0d;
                for (var j = 0; j < xRanges[i].Length; j++)
                {
                    trendVal += coefficients[j] * xRanges[i][xRanges[i].Count() - 1 - j];
                }
                if (columnArray)
                {
                    resultRange.SetValue(i, 0, trendVal + coefficients[coefficients.Length - 1]);
                }
                else
                {
                    resultRange.SetValue(0, i, trendVal + coefficients.Length - 1);
                }
            }
            return resultRange;
        }

        internal static InMemoryRange GetTrendValuesSingle(double[] xRanges, double[] coefficients, bool columnArray)
        {
            var resultRange = (columnArray) ? new InMemoryRange((short)xRanges.Count(), 1) : new InMemoryRange(1, (short)xRanges.Count());
            for (var i = 0; i < xRanges.Count(); i++)
            {
                if (columnArray)
                {
                    resultRange.SetValue(i, 0, xRanges[i] * coefficients[0] + coefficients[1]);
                }
                else
                {
                    resultRange.SetValue(0, i, xRanges[i] * coefficients[0] + coefficients[1]);
                }
                
            }
            return resultRange;
        }
    }
}
