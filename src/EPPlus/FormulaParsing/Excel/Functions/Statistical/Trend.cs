/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the y-values along a linear trend that fits the inputted data. If new_x's is given, it returns the y-values" +
                  "along those x-values.")]
    internal class Trend : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            //Error management
            var constVar = true;
            if (!arguments[0].IsExcelRange) return CompileResult.GetErrorResult(eErrorType.Value);

            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                var argY = arguments[0].ValueAsRangeInfo;
                var argX = arguments[1].ValueAsRangeInfo;

                if (arguments.Count() > 3) constVar = ArgToBool(arguments, 3);

                var linestResult = LinestHelper.ExecuteLinest(argX, argY, constVar, false, false, out eErrorType? error);
                double[] coefficients = new double[linestResult.Size.NumberOfCols];
                for (var i = 0; i < coefficients.Length; i++)
                {
                    coefficients[i] = (double)linestResult.GetValue(0, i);
                }

                bool multipleXranges = false;
                double[][] xRanges;
                if (arguments[2].IsExcelRange)
                {
                    var argNewX = arguments[2].ValueAsRangeInfo;
                    if ((argNewX.Size.NumberOfRows != argY.Size.NumberOfRows && argNewX.Size.NumberOfCols == argY.Size.NumberOfCols)
                    || (argNewX.Size.NumberOfCols != argY.Size.NumberOfCols && argNewX.Size.NumberOfRows == argY.Size.NumberOfRows)) multipleXranges = true;
                    xRanges = LinestHelper.RangeToJaggedDouble(argNewX, argY, constVar, multipleXranges);
                    if (multipleXranges)
                    {
                        return CreateDynamicArrayResult(TrendHelper.GetTrendValuesMultiple(xRanges, coefficients, constVar), DataType.ExcelRange);
                    }
                    else
                    {
                        RangeFlattener.GetNumericPairLists(argNewX, argY, !multipleXranges, out List<double> knownXsList, out List<double> knownYsList);
                        var knownXs = MatrixHelper.ListToArray(knownXsList);
                        return CreateDynamicArrayResult(TrendHelper.GetTrendValuesSingle(knownXs, coefficients), DataType.ExcelRange);
                    }
                }
                else
                {
                    if ((argX.Size.NumberOfRows != argY.Size.NumberOfRows && argX.Size.NumberOfCols == argY.Size.NumberOfCols)
                    || (argX.Size.NumberOfCols != argY.Size.NumberOfCols && argX.Size.NumberOfRows == argY.Size.NumberOfRows)) multipleXranges = true;
                    if (multipleXranges)
                    {
                        xRanges = LinestHelper.RangeToJaggedDouble(argX, argY, constVar, multipleXranges);
                        return CreateDynamicArrayResult(TrendHelper.GetTrendValuesMultiple(xRanges, coefficients, constVar), DataType.ExcelRange);
                    }
                    else
                    {
                        RangeFlattener.GetNumericPairLists(argX, argY, !multipleXranges, out List<double> knownXsList, out List<double> knownYsList);
                        var knownXs = MatrixHelper.ListToArray(knownXsList);
                        return CreateDynamicArrayResult(TrendHelper.GetTrendValuesSingle(knownXs, coefficients), DataType.ExcelRange);
                    }
                }
            }
            else
            {
                //Code for default values here
                var knownYsList = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context, out ExcelErrorValue e1).ToList();
                var knownYs = MatrixHelper.ListToArray(knownYsList);
                var knownXs = LinestHelper.GetDefaultKnownXs(knownYs.Count());
                var linestResult = LinestHelper.LinearRegResult(knownXs, knownYs, constVar, false, false);
                double[] coefficients3 = new double[linestResult.Size.NumberOfCols];
                for (var i = 0; i < coefficients3.Length; i++)
                {
                    coefficients3[i] = (double)linestResult.GetValue(0, i);
                }

                return CreateDynamicArrayResult(TrendHelper.GetTrendValuesSingle(knownXs, coefficients3), DataType.ExcelRange);
            }
        }
    }
}
