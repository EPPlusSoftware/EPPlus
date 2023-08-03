/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  20/06/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the standard error for each predicted y-value for each x-value in the regression.")]
    internal class Steyx : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var knownY = ArgToRangeInfo(arguments, 0);
            var knownX = ArgToRangeInfo(arguments, 1);

            //If an array contains text, logical values or an empty cells, that specific cell is ignored. Cells with value zero are included in the calculations.

            //Difference in data points in the two ranges are not tolerated, and #N/A is returned

            //If the ranges are empty or contains less than three data points, #DIV/0! is returned.

            List<double> yValues = new List<double>();
            List<double> xValues = new List<double>();

            if (knownY.Size.NumberOfRows * knownY.Size.NumberOfCols != knownX.Size.NumberOfRows * knownX.Size.NumberOfCols)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            else if (knownY.Size.NumberOfRows * knownY.Size.NumberOfCols < 3 || knownX.Size.NumberOfRows * knownX.Size.NumberOfCols < 3)
            {
                return CompileResult.GetErrorResult(eErrorType.Div0);
            }

            for (var i = 0; i < knownY.Size.NumberOfRows; i++)
            {
                for (var j = 0; j < knownX.Size.NumberOfCols; j++)
                {
                    var yVal = knownY.GetOffset(i, j);
                    var xVal = knownX.GetOffset(i, j);

                    if (ConvertUtil.IsNumeric(xVal) && ConvertUtil.IsNumeric(yVal))
                    {
                        var doubleyVal = ConvertUtil.GetValueDouble(yVal);
                        var doublexVal = ConvertUtil.GetValueDouble(xVal);

                        yValues.Add(doubleyVal);
                        xValues.Add(doublexVal);
                    }
                   
                }
            }
            var result = SEHelper.GetStandardError(yValues, xValues, false);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
