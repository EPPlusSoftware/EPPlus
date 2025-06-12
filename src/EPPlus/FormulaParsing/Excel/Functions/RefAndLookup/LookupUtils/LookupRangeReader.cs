using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal static class LookupRangeReader
    {
        [DebuggerDisplay("Value: {Value}, Index: {Index}")]
        internal class LookupRangeCell
        {
            public object Value { get; set; }

            public int Index { get; set; }
        }

        public static List<LookupRangeCell> GetLookupRange(IRangeInfo range, out LookupRangeDirection direction)
        {
            var result = new List<LookupRangeCell>();
            if(range.Size.NumberOfRows > range.Size.NumberOfCols)
            {
                direction = LookupRangeDirection.Vertical;
                for(var row = 0; row < range.Size.NumberOfRows; row++)
                {
                    var val = range.GetOffset(row, 0);
                    if (val is ExcelErrorValue) continue;
                    result.Add(new LookupRangeCell { Value = val, Index = row });
                }
            }
            else
            {
                direction = LookupRangeDirection.Horizontal;
                for(var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var val = range.GetOffset(0, col);
                    if (val is ExcelErrorValue) continue;
                    result.Add(new LookupRangeCell { Value = val, Index = col });
                }
            }
            return result;
        }
    }
}
