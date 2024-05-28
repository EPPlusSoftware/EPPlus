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
            format = format.Replace(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".");
            format = format.Replace(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator.Replace((char)160,' '), ","); //Special handling for No-Break Space
            
            var result = context.ExcelDataProvider.GetFormat(value, format);

            return CreateResult(result, DataType.String);
        }
    }
}
