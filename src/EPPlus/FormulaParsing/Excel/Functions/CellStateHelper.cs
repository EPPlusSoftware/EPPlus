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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Utils;
using System;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal static class CellStateHelper
    {
        private static bool IsSubTotal(ICellInfo c, ParsingContext context)
        {
            return (context.Scopes.Current.IsSubtotal && context.SubtotalAddresses.Contains(c.Id));
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, ICellInfo c, ParsingContext context)
        {
            return ShouldIgnore(ignoreHiddenValues, false, c, context);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, bool ignoreNonNumeric, ICellInfo c, ParsingContext context)
        {
            if (ignoreNonNumeric && !ConvertUtil.IsNumericOrDate(c.Value)) return true;
            var hasFilter = false;
            if (context.Parser != null && context.Parser.FilterInfo != null)
            {
                hasFilter = context.Parser.FilterInfo.WorksheetHasFilter(c.WorksheetName);
            }
            return ((ignoreHiddenValues || hasFilter) && c.IsHiddenRow) || IsSubTotal(c, context);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, FunctionArgument arg, ParsingContext context)
        {
            var hasFilter = false;
            if (context.Parser != null && context.Parser.FilterInfo != null && context.Parser.FilterInfo.WorksheetHasFilter(context.Scopes.Current.Address.WorksheetName))
            {
                hasFilter = true;
            }
            return (ignoreHiddenValues || hasFilter) && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell);
        }
    }
}
