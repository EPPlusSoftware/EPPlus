/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/08/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "5.3",
        Description = "Returns the largest numeric value that meets one or more criteria in a range of values.",
        IntroducedInExcelVersion = "2019")]
    internal class MaxIfs : IfsWithMultipleMatchesBase
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var matches = GetMatches("MAXIFS", arguments, out CompileResult errorResult);
            if (errorResult != null)
                return errorResult;
            if (matches.Count() == 0) return CompileResult.ZeroDecimal;
            return CreateResult(matches.Max(), DataType.Decimal);
        }
    }
}
