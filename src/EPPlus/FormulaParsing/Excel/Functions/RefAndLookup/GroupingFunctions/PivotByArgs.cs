using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
