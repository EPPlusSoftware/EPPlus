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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Substitutes all occurrences of a search text string, within an original text string, with the supplied replacement text")]
    internal class Substitute : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            var find = ArgToString(arguments, 1);
            var replaceWith = ArgToString(arguments, 2);
            var result = arguments.Count > 3 ? ReplaceFirst(text, find, replaceWith, ArgToInt(arguments, 3)) : text.Replace(find, replaceWith);
            return CreateResult(result, DataType.String);
        }

        private static string ReplaceFirst(string text, string search, string replace, int instanceNumber)
        {
            int pos = -1;
            for (int i = 0; i < instanceNumber; i++)
            {
                pos = text.IndexOf(search, pos + 1);
                if (pos < 0)
                    break;
            }
            return pos < 0 ? text : text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
    }
}
