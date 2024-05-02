/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "7.2",
        Description = "Get the text before delimiter",
        SupportsArrays = true)]
    internal class TextBefore : ExcelFunctionTextBase
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = ArgToRangeInfo(arguments, 0);
            var text = string.Empty;
            if(range == null)
            {
                text = ArgToString(arguments, 0);
            }
            var delimiters = ArgDelimiterCollectionToString(arguments, 1, out CompileResult error);
            if (error != null) return error;
            var instanceNum = 1;
            var matchMode = "0";
            var matchEnd = 0;
            var ifNotFound = "#N/A";
            var resultString = string.Empty;
            if (arguments.Count > 2)
            {
                instanceNum = ArgToInt(arguments, 2, RoundingMethod.Convert);
                if (instanceNum == 0)
                {
                    instanceNum = 1;
                }
            }
            if (arguments.Count > 3)
            {
                matchMode = ArgToString(arguments, 3);
                if (matchMode == "1")
                {
                    delimiters += delimiters.ToLower() + delimiters.ToUpper();
                }
            }
            if (arguments.Count > 4)
            {
                matchEnd = ArgToInt(arguments, 4, RoundingMethod.Convert);
            }
            if (arguments.Count > 5)
            {
                ifNotFound = ArgToString(arguments, 5);
            }
            int row = range == null ? 0 : range.Address.FromRow;
            int col = range == null ? 0 : range.Address.FromCol;
            var r = range == null ? 1 : (range.Address.ToRow - range.Address.FromRow)+1;
            var c = range == null ? 1 : (range.Address.ToCol - range.Address.FromCol)+1;
            var returnRange = new InMemoryRange(r, (short)c);
            for (int y = 0; y < r; y++)
            {
                for (int x = 0; x < c; x++)
                {
                    text = range == null ? text : range.GetValue(row, col).ToString();
                    col++;
                    int length = 0;
                    int instances = 0;
                    if (instanceNum < 0)
                    {
                        CountBackwards(returnRange, x, y, text, delimiters, instanceNum, matchEnd, ifNotFound, out instances, out length);
                    }
                    else
                    {
                        countForward(returnRange, x, y, text, delimiters, instanceNum, matchEnd, ifNotFound, out instances, out length);
                    }
                }
                row++;
                col = range == null ? 0 : range.Address.FromCol;
            }
            return CreateDynamicArrayResult(returnRange, DataType.ExcelRange);
        }

        private void SetValue(InMemoryRange returnRange, string text, int x, int y, string ifNotFound, int length)
        {
            length = length--;
            if (length <= 0)
            {
                if (ifNotFound != "#N/A")
                {
                    returnRange.SetValue(y, x, ifNotFound);
                    return;
                }
                returnRange.SetValue(y, x, ExcelErrorValue.Create(eErrorType.NA));
                return;
            }
            returnRange.SetValue(y, x, text.Substring(0, length));
        }

        private void countForward(InMemoryRange returnRange, int x, int y, string text, string delimiters, int instanceNum, int matchEnd, string ifNotFound, out int instances, out int length)
        {
            instances = 0;
            length = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char t = text[i];
                if (delimiters.Contains(t))
                {
                    instances++;
                    length = i;
                    if (instances == instanceNum)
                    {
                        break;
                    }
                }
            }
            if (instances < instanceNum && matchEnd == 0)
            {
                if (ifNotFound != "#N/A")
                {
                    returnRange.SetValue(y, x, ifNotFound);
                    return;
                }
                returnRange.SetValue(y, x, ExcelErrorValue.Create(eErrorType.NA));
                return;
            }
            if (matchEnd == 1 && instances - instanceNum == -1)
            {
                returnRange.SetValue(y, x, text);
                return;
            }
            else if (matchEnd == 1 && instances - instanceNum < -1)
            {
                if (ifNotFound != "#N/A")
                {
                    returnRange.SetValue(y, x, ifNotFound);
                    return;
                }
                returnRange.SetValue(y, x, ExcelErrorValue.Create(eErrorType.NA));
                return;
            }
            SetValue(returnRange, text, x, y, ifNotFound, length);
        }

        private void CountBackwards(InMemoryRange returnRange, int x, int y, string text, string delimiters, int instanceNum, int matchEnd, string ifNotFound, out int instances, out int length)
        {
            instances = 0;
            length = 0;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                char t = text[i];
                if (delimiters.Contains(t))
                {
                    instances--;
                    length = i;
                    if (instances == instanceNum) break;
                }
            }
            if (instances > instanceNum && matchEnd == 0)
            {
                if (ifNotFound != "#N/A")
                {
                    returnRange.SetValue(y, x, ifNotFound);
                    return;
                }
                returnRange.SetValue(y, x, ExcelErrorValue.Create(eErrorType.NA));
                return;
            }
            if (matchEnd == 1 && instances - instanceNum == 1)
            {
                returnRange.SetValue(y, x, text);
                return;
            }
            else if (matchEnd == 1 && instances - instanceNum > 1)
            {
                if (ifNotFound != "#N/A")
                {
                    returnRange.SetValue(y, x, ifNotFound);
                    return;
                }
                returnRange.SetValue(y, x, ExcelErrorValue.Create(eErrorType.NA));
                return;
            }
            SetValue(returnRange, text, x, y, ifNotFound, length);
        }
    }
}
