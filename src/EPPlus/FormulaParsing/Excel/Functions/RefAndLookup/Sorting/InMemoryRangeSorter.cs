/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting
{
    internal class InMemoryRangeSorter
    {
        private readonly LookupComparer _comparer = new LookupComparer(LookupMatchMode.ExactMatch);
        public InMemoryRange SortByRow(IRangeInfo sourceRange, int colIndex, int sortOrder)
        {
            var rangeDef = new RangeDefinition(sourceRange.Size.NumberOfRows, sourceRange.Size.NumberOfCols);
            var sortedRange = new InMemoryRange(rangeDef);
            var columns = new List<SortedColOrRow>();
            for(var col = 0; col < rangeDef.NumberOfCols; col++)
            {
                var rows = new SortedColOrRow();
                for(var row = 0; row < rangeDef.NumberOfRows; row++)
                {
                    var v = sourceRange.GetOffset(row, col);
                    var si = new InMemoryRangeSortItem(v, row);
                    rows.AddItem(row, si);
                }
                columns.Add(rows);
            }
            var colIx = colIndex - 1;
            var colToSortList = columns[colIx].ToList();
            colToSortList.Sort((a, b) => _comparer.Compare(a.Value, b.Value) * sortOrder);
            for (var row = 0; row < colToSortList.Count; row++)
            {
                var sortedColItem = colToSortList[row];
                sortedRange.SetValue(row, colIx, sortedColItem.Value);
                for (var col = 0; col < columns.Count; col++)
                {
                    if (col == colIx) continue;
                    var colItem = columns[col].GetByOriginalIndex(sortedColItem.OriginalIndex);
                    sortedRange.SetValue(row, col, colItem.Value);
                    
                }
            }
            return sortedRange;
        }

        public InMemoryRange SortByCol(IRangeInfo sourceRange, int rowIndex, int sortOrder)
        {
            var rangeDef = new RangeDefinition(sourceRange.Size.NumberOfRows, sourceRange.Size.NumberOfCols);
            var sortedRange = new InMemoryRange(rangeDef);
            var rows = new List<SortedColOrRow>();
            for (var row = 0; row < rangeDef.NumberOfRows; row++)
            {
                var cols = new SortedColOrRow();
                for (var col = 0; col < rangeDef.NumberOfCols; col++)
                {
                    var v = sourceRange.GetOffset(row, col);
                    var si = new InMemoryRangeSortItem(v, col);
                    cols.AddItem(col, si);
                }
                rows.Add(cols);
            }
            var rowIx = rowIndex - 1;
            var rowToSortList = rows[rowIx].ToList();
            rowToSortList.Sort((a, b) => _comparer.Compare(a.Value, b.Value) * sortOrder);
            for (var col = 0; col < rowToSortList.Count; col++)
            {
                var sortedRowItem = rowToSortList[col];
                sortedRange.SetValue(rowIx, col, sortedRowItem.Value);
                for (var row = 0; row < rows.Count; row++)
                {
                    if (row == rowIx) continue;
                    var colItem = rows[row].GetByOriginalIndex(sortedRowItem.OriginalIndex);
                    sortedRange.SetValue(row, col, colItem.Value);

                }
            }
            return sortedRange;
        }
    }
}
