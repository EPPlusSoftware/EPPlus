/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 05/07/2023         EPPlus Software AB         Implemented function
*************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class GrowthHelper
    {
        internal static InMemoryRange GetGrowthValuesMultiple(double[][] xRanges, double[] coefficients, bool constVar, bool columnArray)
        {
            var resultRange = (columnArray) ? new InMemoryRange((short)xRanges.Count(), 1) : new InMemoryRange(1, (short)xRanges.Count());
            //If const, we remove the column of ones.
            List<double> removeColumn = [xRanges[0].Count() - 1];
            if (constVar) xRanges = MatrixHelper.RemoveColumns(xRanges, removeColumn);
            for (var i = 0; i < xRanges.Length; i++)
            {
                //Formula for the multiple variable estimation is y = b * m1^x1 * m2^x2 * m3^x3 * ... * mn^xn
                var growthVal = 1d;
                for (var j = 0; j < xRanges[i].Length; j++)
                {
                    growthVal *= Math.Pow(coefficients[j], xRanges[i][xRanges[i].Count() - 1 - j]);
                }
                if (columnArray)
                {
                    resultRange.SetValue(i, 0, growthVal * coefficients[coefficients.Length - 1]);
                }
                else
                {
                    resultRange.SetValue(0, i, growthVal * coefficients[coefficients.Length - 1]);
                }
            }
            return resultRange;
        }

        internal static InMemoryRange GetGrowthValuesSingle(double[] xRanges, double[] coefficients, bool columnArray)
        {
            var resultRange = (columnArray) ? new InMemoryRange((short)xRanges.Count(), 1) : new InMemoryRange(1, (short)xRanges.Count());
            for (var i = 0; i < xRanges.Count(); i++)
            {
                if (columnArray)
                {
                    resultRange.SetValue(i, 0, coefficients[1] * Math.Pow(coefficients[0], xRanges[i])); //Formula for the estimation is y = b * m^x
                }
                else
                {
                    resultRange.SetValue(0, i, coefficients[1] * Math.Pow(coefficients[0], xRanges[i]));
                }

            }
            return resultRange;
        }
    }
}