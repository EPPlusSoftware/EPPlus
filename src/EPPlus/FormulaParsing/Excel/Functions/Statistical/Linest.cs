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
       Description = "The LINEST function calculates a regressional line that fits your data. It also calculates additional statistics." +
                     "It can handle several x-variables and perform multiple regression analysis.")]
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
                else if (multipleXranges && xRows != yRows)
                {
                    rowArray = true;
                    for (var i = 0; i < xRows; i++)
                    {
                        xRanges.Add(i, new List<double>());
                    }

                    var rowCount = 0;
                    var listCount = 0;
                    while (rowCount < xRows)
                    {
                        rowCount += 1;
                        while (listCount < (xColumns * rowCount))
                        {
                            xRanges[rowCount - 1].Add(knownXs[listCount]); //Write test for this
                            listCount += 1;
                        }
                    }
                }

                if (arguments.Count() > 2 && arguments[2].DataType != DataType.Empty) constVar = ArgToBool(arguments, 2); //Need to change this
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

                List<List<double>> xRangeList = new List<List<double>>();
                if (columnArray)
                {
                    for (var i = 0; i < xColumns; i++)
                    {
                        xRangeList.Add(xRanges[i]);
                    }
                }
                else if (rowArray)
                {
                    for (var i = 0; i < xRows; ++i)
                    {
                        xRangeList.Add(xRanges[i]);
                    }
                }
                if (multipleXranges)
                {
                    var resultRangeX = LinestHelper.CalculateMultipleXRanges(knownYs, xRangeList, constVar, stats);
                    return CreateResult(resultRangeX, DataType.ExcelRange);
                }
                else
                {
                    var resultRange = LinestHelper.CalculateResult(knownYs, knownXs, constVar, stats);
                    return CreateResult(resultRange, DataType.ExcelRange);
                }
            }
            else
            {
                var knownYs = ArgsToDoubleEnumerable(new List<FunctionArgument> {  arguments[0] }, context).Select(x => (double)x).ToList();
                var knownXs = GetDefaultKnownXs(knownYs.Count());

                if (arguments.Count() > 2) constVar = ArgToBool(arguments, 2);
                if (arguments.Count() > 3) stats = ArgToBool(arguments, 3);

                var resultRange = LinestHelper.CalculateResult(knownYs, knownXs, constVar, stats); //change here so that multiple x is possible
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
    }
}
