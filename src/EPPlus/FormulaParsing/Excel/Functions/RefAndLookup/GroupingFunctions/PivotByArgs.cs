/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/4/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions
{
    internal class PivotByArgs : GroupByBaseArgs
    {
        public IRangeInfo ColFields { get; set; }
        public int RowTotalDepth { get; set; } = 0;
        public int[] RowSortOrders { get; set; } = new[] { 1 };
        public int ColTotalDepth { get; set; } = 0;
        public int[] ColSortOrders { get; set; } = new[] { 1 };
        public RelativeTo RelativeTo { get; set; } = RelativeTo.ColumnTotals; // Default
    }
}
