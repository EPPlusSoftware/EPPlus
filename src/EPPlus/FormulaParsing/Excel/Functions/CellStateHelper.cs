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
        private static bool ShouldIgnoreNestedSubtotal(bool ignoreNestedSubtotalAndAggregate, ulong cellId, ParsingContext context)
        {
            if (!ignoreNestedSubtotalAndAggregate) return false;
            return context.SubtotalAddresses.Contains(cellId);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, bool ignoreNestedSubtotalAndAggregate, ICellInfo c, ParsingContext context)
        {
            return ShouldIgnore(ignoreHiddenValues, false, c, context, ignoreNestedSubtotalAndAggregate);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, ICellInfo c, ParsingContext context)
        {
            return ShouldIgnore(ignoreHiddenValues, false, c, context);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, bool ignoreNonNumeric, ICellInfo c, ParsingContext context, bool ignoreNestedSubtotalAndAggregate = true)
        {
            if(c.Address==null) return false;
            if (ignoreNonNumeric && !ConvertUtil.IsNumericOrDate(c.Value)) return true;
            var filterExists = false;
            if (context.HiddenCellBehaviour == HiddenCellHandlingCategory.Subtotal 
                && context.Parser != null 
                && context.Parser.FilterInfo != null)
            {
                filterExists = context.Parser.FilterInfo.CellIsCoveredByFilter(context.Package.Workbook.Worksheets[c.WorksheetName].IndexInList, c);
            }
            return ((ignoreHiddenValues || filterExists) && c.IsHiddenRow) || ShouldIgnoreNestedSubtotal(ignoreNestedSubtotalAndAggregate, c.Id, context);
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, bool ignoreNestedSubtotalAndAggregate, FunctionArgument arg, ParsingContext context)
        {
            var filterExists = false;
            if (context.HiddenCellBehaviour == HiddenCellHandlingCategory.Subtotal
                && context.Parser != null 
                && context.Parser.FilterInfo != null 
                && context.Parser.FilterInfo.CellIsCoveredByFilter(context.CurrentCell.WorksheetIx, arg))
            {
                filterExists = true;
            }
            var include = true;
            if(ignoreNestedSubtotalAndAggregate && arg.Address != null)
            {
                var cellId = arg.Address.GetTopLeftCellId();
                include = !ShouldIgnoreNestedSubtotal(ignoreNestedSubtotalAndAggregate, cellId, context);
            }
            return (ignoreHiddenValues || filterExists) && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell) && include;
        }
    }
}
