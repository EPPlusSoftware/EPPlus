/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Statistical,
       EPPlusVersion = "7.0",
       Description = "The LINEST function calculates...")]
    internal class Linest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);
            var constVar = true;
            var stats = false;
            if (arguments.Count() > 1)
            {
                var rangeX = arguments[1].ValueAsRangeInfo;
                var rangeY = arguments[0].ValueAsRangeInfo;
                var xColumns = rangeX.Size.NumberOfCols;
                var yColumns = rangeY.Size.NumberOfCols;
                var xRows = rangeX.Size.NumberOfRows;
                var yRows = rangeY.Size.NumberOfRows;
                if (xRows != yRows || xColumns != yColumns) return CreateResult(eErrorType.Ref);
                RangeFlattener.GetNumericPairLists(rangeX, rangeY, true, out List<double> knownXs, out List<double> knownYs);
                if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2);
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);
                var resultRange = CalculateResult(knownYs, knownXs, constVar, stats);
                return CreateResult(resultRange, DataType.ExcelRange);
            }
            else
            {
                var knownYs = ArgsToDoubleEnumerable(new List<FunctionArgument> {  arguments[0] }, context).Select(x => (double)x).ToList();
                var knownXs = GetDefaultKnownXs(knownYs.Count());
                if (arguments.Count() > 2) constVar = ArgToBool(arguments, 2);
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);
                var resultRange = CalculateResult(knownYs, knownXs, constVar, stats);
                return CreateResult(resultRange, DataType.ExcelRange);
            }

        }
        private List<double> GetDefaultKnownXs(int count)
        {
            List<double> result = new List<double>();
            for (int i = 1; i <= count; i++)
            {
                result.Add(i);
            }
            return result;
        }

        private static double StandardDeviation(List<double> values)
        {
            //Returns the standard deviation of a list

            var std = 0d;
            var mean = values.Average();

            for (var i = 0; i < values.Count; i++)
            {
                std += Math.Pow(values[i] - mean, 2);
            }

            std = Math.Sqrt(std / (values.Count() - 1));

            return std;
        }

        private static double GetCoefficientOfDetermination(List<double> estimated, List<double> actual)
        {
            var nominator = 0d;
            var denominator = 0d;
            var estimatedMean = estimated.Average();
            var actualMean = actual.Average();
            for (var i = 0; i < estimated.Count; i++)
            {
                nominator += (estimated[i] - estimatedMean) * (actual[i] - actualMean);
                denominator += Math.Pow(estimated[i] - actualMean, 2) * Math.Pow(actual[i] - actualMean, 2);
            }

            return Math.Pow(nominator / denominator, 2);
        }

        private InMemoryRange CalculateResult(List <double> knownYs, List <double> knownXs, bool constVar, bool stats)
        {
            var averageY = knownYs.Average();
            var averageX = knownXs.Average();
            
            double nominator = 0d;
            double denominator = 0d;
            double xDiff = 0d;
            double yDiff = 0d;
            double estimatedDiff = 0d;
            double ssr = 0d;
            double sst = 0d;
            var df = 0d;
            var v1 = 0d;
            var v2 = 0d;
            var fStatistics = 0d;
            var f = 0d;

            for (var i = 0; i < knownYs.Count; i++)
            {
                var y = knownYs[i];
                var x = knownXs[i];

                if (constVar)
                {
                    nominator += (x - averageX) * (y - averageY);
                    denominator += (x - averageX) * (x - averageX);
                }
                else
                {
                    nominator += x * y;
                    denominator += Math.Pow(x, 2);
                }

            }

            var m = nominator / denominator;
            var b = (constVar) ? averageY - (m * averageX) : 0d;

            if (stats)
            {
                for (var i = 0; i < knownXs.Count(); i++)
                {
                    var x = knownXs[i];
                    var y = knownYs[i];
                    var estimatedY = m * x + b;

                    estimatedDiff += Math.Pow(y - estimatedY, 2);
                    xDiff += Math.Pow(x - averageX, 2);
                    yDiff += Math.Pow(y - estimatedY, 2);
                    ssr += Math.Pow(estimatedY - averageY, 2);
                    sst += Math.Pow(y - averageY, 2);
                }

                var errorVariance = yDiff / (knownXs.Count() - 2);
                var standardErrorM = Math.Sqrt(1d / (knownXs.Count() - 2d) * estimatedDiff / xDiff);
                var standardErrorB = Math.Sqrt(errorVariance) * Math.Sqrt(1d / knownXs.Count() + Math.Pow(averageX, 2) / xDiff);
                var ssreg = ssr;
                var ssresid = yDiff;
                var rSquared = ssr / sst;
                var standardErrorEstimateY = SEHelper.GetStandardError(knownXs, knownYs);

                if (constVar)
                {
                    df = knownXs.Count() - 2; //Need to review this
                    v1 = knownXs.Count() - df - 1;
                    v2 = df;
                    fStatistics = (ssr / v1) / (yDiff / v2);
                }
                else
                {
                    df = knownXs.Count() - 1; //Need to review this
                    v1 = knownXs.Count() - df;
                    v2 = df;
                    fStatistics = ssr / yDiff;
                }

                var resultRangeStats = new InMemoryRange(5, 2);
                resultRangeStats.SetValue(0, 0, m);
                resultRangeStats.SetValue(0, 1, b);
                resultRangeStats.SetValue(1, 0, standardErrorM);
                resultRangeStats.SetValue(1, 1, standardErrorB);
                resultRangeStats.SetValue(2, 0, rSquared);
                resultRangeStats.SetValue(2, 1, standardErrorEstimateY);
                resultRangeStats.SetValue(3, 0, fStatistics);
                resultRangeStats.SetValue(3, 1, df);
                resultRangeStats.SetValue(4, 0, ssreg);
                resultRangeStats.SetValue(4, 1, ssresid);
                return resultRangeStats;
            }

            var resultRangeNormal = new InMemoryRange(1, 2);
            resultRangeNormal.SetValue(0, 0, m);
            resultRangeNormal.SetValue(0, 1, b);
            return resultRangeNormal;


        }
    }
}
