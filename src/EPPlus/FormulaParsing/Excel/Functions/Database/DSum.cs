/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Database,
        EPPlusVersion = "4",
        Description = "Calculates the sum of values in a field of a list or database, that satisfy specified conditions")]
    internal class Dsum : DatabaseFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var values = GetMatchingValues(arguments, context);
            if (!values.Any()) return CreateResult(0d, DataType.Integer);
            return CreateResult(values.Sum(), DataType.Integer);
        }
    }
}
