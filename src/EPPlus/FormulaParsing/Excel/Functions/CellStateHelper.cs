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
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal static class CellStateHelper
    {
        private static bool IsSubTotal(ExcelDataProvider.ICellInfo c)
        {
            var tokens = c.Tokens;
            if (tokens == null) return false;
            return c.Tokens.Any(token => 
                token.TokenTypeIsSet(LexicalAnalysis.TokenType.Function) 
                && (token.Value.Equals("SUBTOTAL", StringComparison.OrdinalIgnoreCase) || token.Value.Equals("AGGREGATE", StringComparison.OrdinalIgnoreCase))
                );
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, ExcelDataProvider.ICellInfo c, ParsingContext context)
        {
            return (ignoreHiddenValues && c.IsHiddenRow) || (context.Scopes.Current.IsSubtotal && IsSubTotal(c));
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, FunctionArgument arg, ParsingContext context)
        {
            return (ignoreHiddenValues && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }
    }
}
