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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
       Description = "The LINEST function calculates a regressional line that fits your data. It also calculates additional statistics.")]
    internal class Linest : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            //X can have more than one vector corresponding to each y-value
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);
            var constVar = true;
            var stats = false;
            bool multipleXranges = false;
            bool columnArray = false;
            bool rowArray = false;
            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                var rangeX = arguments[1].ValueAsRangeInfo;
                var rangeY = arguments[0].ValueAsRangeInfo;
                var xColumns = rangeX.Size.NumberOfCols;
                var yColumns = rangeY.Size.NumberOfCols;
                var xRows = rangeX.Size.NumberOfRows;
                var yRows = rangeY.Size.NumberOfRows;

                if ((xRows != yRows && xColumns == yColumns)
                    || (xColumns != yColumns && xRows == yRows))
                {
                    multipleXranges = true;
                }
                else
                {
                    if (xRows != yRows || xColumns != yColumns) return CreateResult(eErrorType.Ref);
                }

                RangeFlattener.GetNumericPairLists(rangeX, rangeY, !multipleXranges, out List<double> knownXs, out List<double> knownYs);
                SortedDictionary<int, List<double>> xRanges = new SortedDictionary<int, List<double>>();

                if (multipleXranges && xColumns != yColumns)
                {
                    columnArray = true;
                    for (var i = 0; i < xColumns; i++)
                    {
                        xRanges.Add(i, new List<double>());
                    }

                    var colCount = -1;

                    while (colCount < (xColumns - 1))
                    {
                        colCount += 1;
                        var listCount = colCount;
                        while (listCount < knownXs.Count())
                        {
                            xRanges[colCount].Add(knownXs[listCount]);
                            listCount += xColumns;
                        }
                    }
                }
                else if (multipleXranges && xRows != yRows) //This is wrong!!! goes through all columns so rows is "easier"
                                                            //Change later
                {
                    rowArray = true;
                    for (var i = 0; i < xRows; i++)
                    {
                        xRanges.Add(i, new List<double>());
                    }

                    var rowCount = -1;

                    while (rowCount < (xRows - 1))
                    {
                        rowCount += 1;
                        var listCount = rowCount;
                        while (listCount < knownXs.Count())
                        {
                            xRanges[rowCount].Add(knownXs[listCount]);
                            listCount += xRows;
                        }
                    }
                }

                if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2);
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);
                if (columnArray)
                {
                    List<List<double>> xRangeList = new List<List<double>>();

                    for (var i = 0; i < xColumns; i++)
                    {
                        xRangeList.Add(xRanges[i]);
                    }
                }
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

                    if (constVar)
                    {
                        estimatedDiff += Math.Pow(y - estimatedY, 2);
                        xDiff += Math.Pow(x - averageX, 2);
                        yDiff += Math.Pow(y - estimatedY, 2);
                        ssr += Math.Pow(estimatedY - averageY, 2);
                        sst += Math.Pow(y - averageY, 2);
                    }
                    else
                    {
                        estimatedDiff += Math.Pow(y - estimatedY, 2);
                        xDiff += Math.Pow(x, 2);
                        yDiff = Math.Pow(y - estimatedY, 2);
                        ssr += Math.Pow(estimatedY, 2);
                        sst += Math.Pow(y, 2);
                    }

                }

                var errorVariance = yDiff / (knownXs.Count() - 2);
                if (!constVar) errorVariance = yDiff / (knownXs.Count() - 1);

                var standardErrorM = (constVar) ? Math.Sqrt(1d / (knownXs.Count() - 2d) * estimatedDiff / xDiff) : 
                                                  Math.Sqrt(1d / (knownXs.Count() - 1d) * estimatedDiff / xDiff) ;

                object standardErrorB = Math.Sqrt(errorVariance) * Math.Sqrt(1d / knownXs.Count() + Math.Pow(averageX, 2) / xDiff);
                if (!constVar) standardErrorB = ExcelErrorValue.Create(eErrorType.NA);

                var rSquared = ssr / sst;
                var standardErrorEstimateY = (!constVar) ? SEHelper.GetStandardError(knownXs, knownYs, true) :
                                                          SEHelper.GetStandardError(knownXs, knownYs, false) ;
                var ssreg = ssr;
                var ssresid = (constVar) ? yDiff : (sst - ssr);

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
                    fStatistics = ssr / (ssresid / (knownXs.Count() - 1));
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
