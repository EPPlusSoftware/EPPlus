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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Returns a specified number of characters from the end of a supplied text string",
        SupportsArrays = true)]
    internal class Right : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var str = ArgToString(arguments, 0);
            if (str == null)
                str = string.Empty;
            var length = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (length < 0) return CompileResult.GetErrorResult(eErrorType.Value);
            var startIx = str.Length - length;
            if (startIx < 0)
                startIx = 0;

            if (context.Configuration.EnableUnicodeAwareStringOperations)
            {
                int unicodeLength = 0;
                for (int i = str.Length - 1; i >= 0; i--)
                {
                    char c = str[i];

                    // Handle surrogate pairs correctly
                    if (char.IsLowSurrogate(c) && i > 0 && char.IsHighSurrogate(str[i - 1]))
                    {
                        i--; // Move past the high surrogate
                    }

                    unicodeLength++;

                    if (unicodeLength == length)
                    {
                        startIx = i; // Set the correct start position
                        break;
                    }
                }
                var res = str.UnitcodeSubstring(startIx, str.Length - startIx);
                return CreateResult(res, DataType.String);
            }
            return CreateResult(str.Substring(startIx, str.Length - startIx), DataType.String);
        }
    }
}
