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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the y-values along an exponential curve that best fits the inputted data. If new_x's is given, it returns the y-values" +
                  "along those x-values. Growth can also find best fitting curve for a model with multiple predictor variables.")]
    internal class Growth : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            //Error management
            bool constVar = true; //default value
            bool columnArray = false;
            bool multipleXranges = false;
            IRangeInfo argY = arguments[0].ValueAsRangeInfo;
            IRangeInfo argNewX;
            IRangeInfo argX;
            IRangeInfo logestResult;
            if (argY.Size.NumberOfCols == 1) columnArray = true;

            if (arguments.Count() > 1 && arguments[1].IsExcelRange)
            {
                argX = arguments[1].ValueAsRangeInfo;
            }
            else
            {
                //Code for default values here
                var xVals = LinestHelper.GetDefaultKnownXs(argY.Count());
                if (arguments.Count() > 3) constVar = ArgToBool(arguments, 3);
                argX = LinestHelper.GetDefaultKnownXsRange(argY);
                logestResult = LinestHelper.ExecuteLinest(argX, argY, constVar, false, true, out eErrorType? defaultError);
                if (defaultError != null) return CreateResult(defaultError.Value);
                double[] defaultCoefficients = new double[logestResult.Size.NumberOfCols];
                for (var i = 0; i < defaultCoefficients.Length; i++)
                {
                    defaultCoefficients[i] = (double)logestResult.GetValue(0, i);
                }
                return CreateDynamicArrayResult(GrowthHelper.GetGrowthValuesSingle(xVals, defaultCoefficients, columnArray), DataType.ExcelRange);
            }

            if (arguments.Count() > 3) constVar = ArgToBool(arguments, 3);

            //Get the line coefficient(s)
            if ((argX.Size.NumberOfRows != argY.Size.NumberOfRows && argX.Size.NumberOfCols == argY.Size.NumberOfCols)
            || (argX.Size.NumberOfCols != argY.Size.NumberOfCols && argX.Size.NumberOfRows == argY.Size.NumberOfRows)) multipleXranges = true;
            if (multipleXranges && argX.Size.NumberOfCols != argY.Size.NumberOfCols) columnArray = true;

            logestResult = LinestHelper.ExecuteLinest(argX, argY, constVar, false, true, out eErrorType? error);
            if (error != null) return CreateResult(error.Value);
            double[] coefficients = new double[logestResult.Size.NumberOfCols];
            for (var i = 0; i < coefficients.Length; i++)
            {
                coefficients[i] = (double)logestResult.GetValue(0, i);
            }

            //If newXs is given:
            if (arguments[2].IsExcelRange)
            {
                argNewX = arguments[2].ValueAsRangeInfo;
                if (multipleXranges)
                {
                    //knownXs and NewXs must have the same amount of variables, but doesnt have to have the same amount of observations/samples
                    if (columnArray && argNewX.Size.NumberOfCols != argX.Size.NumberOfCols) return CompileResult.GetErrorResult(eErrorType.Ref);
                    if (!columnArray && argNewX.Size.NumberOfRows != argX.Size.NumberOfRows) return CompileResult.GetErrorResult(eErrorType.Ref);

                    var xRanges = LinestHelper.GetRangeAsJaggedDouble(argNewX, argY, constVar, multipleXranges);
                    return CreateDynamicArrayResult(GrowthHelper.GetGrowthValuesMultiple(xRanges, coefficients, constVar, columnArray), DataType.ExcelRange);
                }
                else
                {
                    RangeFlattener.GetNumericPairLists(argNewX, argY, !multipleXranges, out List<double> xVals, out List<double> yVals);
                    var xValsArray = MatrixHelper.ListToArray(xVals);
                    if (argNewX.Size.NumberOfCols == 1) columnArray = true;
                    return CreateDynamicArrayResult(GrowthHelper.GetGrowthValuesSingle(xValsArray, coefficients, columnArray), DataType.ExcelRange);
                }
            }

            //If newXs is omitted:
            if (multipleXranges)
            {
                var xRanges = LinestHelper.GetRangeAsJaggedDouble(argX, argY, constVar, multipleXranges);
                return CreateDynamicArrayResult(GrowthHelper.GetGrowthValuesMultiple(xRanges, coefficients, constVar, columnArray), DataType.ExcelRange);
            }

            //Return for single variable case
            RangeFlattener.GetNumericPairLists(argX, argY, !multipleXranges, out List<double> knownXsList, out List<double> knownYsList);
            var knownXs = MatrixHelper.ListToArray(knownXsList);
            return CreateDynamicArrayResult(GrowthHelper.GetGrowthValuesSingle(knownXs, coefficients, columnArray), DataType.ExcelRange);

        }
    }
}