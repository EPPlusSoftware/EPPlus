﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Returns the array in a single column.",
        SupportsArrays = true)]
    internal class ToCol : ToRowColBase
    {
        public override string NamespacePrefix => "_xlfn.";
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.First();
            var ignore = 0;
            if (arguments.Count() > 1 && arguments.ElementAt(1).Value != null)
            {
                ignore = ArgToInt(arguments, 1);
                if (ignore < 0 || ignore > 4)
                {
                    return CompileResult.GetErrorResult(eErrorType.Value);
                }
            }
            var scanByColumn = false;
            if (arguments.Count() > 2 && arguments.ElementAt(2).Value != null)
            {
                scanByColumn = ArgToBool(arguments, 2);
            }
            if(firstArg.IsExcelRange)
            {
                var result = GetItemsFromRange(firstArg.ValueAsRangeInfo, ignore, scanByColumn);
                var resultRange = new InMemoryRange(new RangeDefinition(result.Count, 1));
                var row = 0;
                foreach (var val in result)
                {
                    resultRange.SetValue(row++, 0, val);
                }
                return CreateResult(resultRange, DataType.ExcelRange);
            }
            return CompileResultFactory.Create(firstArg.Value);
        }
    }
}
