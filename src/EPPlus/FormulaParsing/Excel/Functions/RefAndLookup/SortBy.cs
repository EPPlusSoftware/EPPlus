/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Sorts the contents of a range or array based on the values in a corresponding range or array.",
        SupportsArrays = true)]
    internal class SortBy : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var range = ArgToRangeInfo(arguments, 0);
            var nArgs = arguments.Count();
            for(var x = 1; x < nArgs; x+=2)
            {
                var byRange = ArgToRangeInfo(arguments, x);
                if (byRange.Size.NumberOfCols != range.Size.NumberOfCols && byRange.Size.NumberOfRows != range.Size.NumberOfRows)
                {
                    return CreateResult(eErrorType.Value);
                }
                if(byRange.Size.NumberOfRows > 1 && byRange.Size.NumberOfCols > 1)
                {
                    return CreateResult(eErrorType.Value);
                }
                var sortOrder = 1;
                if(x +1 < nArgs)
                {
                    sortOrder = ArgToInt(arguments, x + 1);

                }
            }

            throw new NotImplementedException();
        }
    }
}
