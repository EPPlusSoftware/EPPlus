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
        Description = "Converts a supplied value into text, using a user-specified format")]
    public class Text : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var value = arguments[0].ValueFirst;
            var format = ArgToString(arguments, 1);
            format = ChangeFormatToEnglishFormat(format);
            var result = context.ExcelDataProvider.GetFormat(value, format);

            return CreateResult(result, DataType.String);
        }

        private static string ChangeFormatToEnglishFormat(string format)
        {
            var decSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var groupSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            if (ExcelWorkbook.Culture != null)
            {
                decSep = ExcelWorkbook.Culture.NumberFormat.NumberDecimalSeparator;
                groupSep = ExcelWorkbook.Culture.NumberFormat.NumberGroupSeparator;
            }

            //DecimalSeparator and GroupSeperator can switch e.g. Culture is German. Using only replace would result in deleting one of the separators e.g
            //"###.###,###".Replace(".",",") => "###,###,###".Replace(",",".") => "###.###.###"
            //the correct replacement in this example would be "###,###.###"

            var groupSepSplit = format.Split(groupSep[0]);
            var englishFormat = "";
            groupSepSplit.ToList().ForEach(g => englishFormat += ("," + g.Replace(decSep, ".")));
            if (!format.StartsWith(groupSep))
            {
                englishFormat = englishFormat.Substring(1);
            }

            return englishFormat;
        }
    }
}
