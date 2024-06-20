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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    /// <summary>
    /// The Text
    /// </summary>
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Converts a supplied value into text, using a user-specified format")]
    public class Text : ExcelFunction
    {
        /// <summary>
        /// Minimum arguments
        /// </summary>
        public override int ArgumentMinLength => 1;
        /// <summary>
        /// Execute function
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var value = arguments[0].ValueFirst;
            var format = ArgToString(arguments, 1);
            var invariantFormat = GetInvariantFormat(format);

            var result = context.ExcelDataProvider.GetFormat(value, invariantFormat);

            return CreateResult(result, DataType.String);
        }

        private static string GetInvariantFormat(string format)
        {
            var nds = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var ds = nds[0];
            var ngs = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var gs = ngs[0];
            var sb = new StringBuilder();
            bool isInString = false;
            for (var i = 0; i < format.Length; i++)
            {
                var c = format[i];
                if (isInString == false)
                {
                    if (c == ds && (nds.Length<=1 || format.Substring(i, nds.Length).Equals(nds)))
                    {
                        sb.Append(".");
                    }
                    else if (c == gs && (ngs.Length <= 1 || format.Substring(i, ngs.Length).Equals(ngs)))
                    {
                        sb.Append(",");
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    sb.Append(c);
                }
                if (c == '\'')
                {
                    isInString = !isInString; ;
                }
            }
            format = sb.ToString();
            return format;
        }
    }
}
