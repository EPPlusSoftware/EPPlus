﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/23/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "5.2",
    Description = "Returns the Mode (the most frequently occurring value) of a list of supplied numbers ")]
    internal class Mode : HiddenValuesHandlingFunction
    {
        public Mode()
        {
            IgnoreErrors = false;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count() > 255) return CompileResult.GetErrorResult(eErrorType.Value);
            var numbers = ArgsToDoubleEnumerable(arguments, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var counts = new Dictionary<double, int>();
            double? maxCountValue = default;
            foreach(var number in numbers)
            {
                if (!counts.ContainsKey(number))
                    counts[number] = 1;
                else
                    counts[number] = counts[number] + 1;
                if(counts[number] > 1 && (!maxCountValue.HasValue || counts[number] >= counts[maxCountValue.Value]))
                {
                    if(!maxCountValue.HasValue)
                    {
                        maxCountValue = number;
                    }
                    else if (counts[number] == counts[maxCountValue.Value] && number < maxCountValue)
                        maxCountValue = number;
                    else if (counts[number] > counts[maxCountValue.Value])
                        maxCountValue = number;
                }
            }
            if (!maxCountValue.HasValue) return CreateResult(eErrorType.Num);
            return CreateResult(maxCountValue.Value, DataType.Decimal);
        }
    }
}
