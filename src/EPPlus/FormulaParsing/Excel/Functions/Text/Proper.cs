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
using System.Globalization;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Converts all characters in a supplied text string to proper case (i.e. letters that do not follow another letter are upper case and all other characters are lower case)")]
    internal class Proper : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            if (text == null)
                text = string.Empty;
            text = text.ToLower(CultureInfo.InvariantCulture);
            var sb = new StringBuilder();
            var previousChar = '.';
            foreach (var ch in text)
            {
                if (!char.IsLetter(previousChar))
                {
                    sb.Append(Utils.ConvertUtil._invariantTextInfo.ToUpper(ch.ToString()));
                }
                else
                {
                    sb.Append(ch);
                }
                previousChar = ch;
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
