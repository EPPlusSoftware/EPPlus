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
        public InMemoryRange SortByCol(IRangeInfo sourceRange, List<int> rowIndexes, int sortOrder, ParsingContext context)
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
            //var colIx = colIndex - 1;
            //var colToSortList = columns[colIx].ToList();
            //var sortedList = colToSortList.Where(x => x.Value != null).ToList();
            //sortedList.Sort((a, b) => _comparer.Compare(a.Value, b.Value, sortOrder));
            //var nullValues = colToSortList.Where(x => x.Value == null);
            //sortedList.AddRange(nullValues);
            //for (var row = 0; row < sortedList.Count; row++)
            //{
            //    var sortedColItem = sortedList[row];
            //    sortedRange.SetValue(row, colIx, sortedColItem.Value);
            //    for (var col = 0; col < columns.Count; col++)
            //    {
            //        if (col == colIx) continue;
            //        var colItem = columns[col].GetByOriginalIndex(sortedColItem.OriginalIndex);
            //        sortedRange.SetValue(row, col, colItem.Value);

            //    }
            //}
            columns.Sort((a, b) =>
            {
                foreach (var rowIx in rowIndexes)
                {
                    var aColVal = a.GetByOriginalIndex(rowIx - 1);
                    var bColVal = b.GetByOriginalIndex(rowIx - 1);
                    if (aColVal.Value == null)
                    {
                        if (bColVal.Value == null) return 0;
                        return 1;
                    }
                    else if (bColVal.Value == null)
                    {
                        return -1;
                    }
                    var res = _comparer.Compare(aColVal.Value, bColVal.Value, sortOrder, context);
                    if (res != 0) return res;
                }
                return 0;
            });
            var colIx = 0;
            foreach (var col in columns)
            {
                var cellIx = 0;
                foreach (var cell in col.ToList())
                {
                    sortedRange.SetValue(cellIx, colIx, cell.Value);
                    cellIx++;
                }
                colIx++;
            }
            return sortedRange;
        }

        public InMemoryRange SortByRow(IRangeInfo sourceRange, List<int> colIndexes, int sortOrder, ParsingContext context)
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
            //var rowIx = rowIndex - 1;
            //var rowToSortList = rows[rowIx].ToList();
            //rowToSortList.Sort((a, b) => _comparer.Compare(a.Value, b.Value, sortOrder));
            //for (var col = 0; col < rowToSortList.Count; col++)
            //{
            //    var sortedRowItem = rowToSortList[col];
            //    sortedRange.SetValue(rowIx, col, sortedRowItem.Value);
            //    for (var row = 0; row < rows.Count; row++)
            //    {
            //        if (row == rowIx) continue;
            //        var colItem = rows[row].GetByOriginalIndex(sortedRowItem.OriginalIndex);
            //        sortedRange.SetValue(row, col, colItem.Value);

            //    }
            //}

            rows.Sort((a, b) => 
            { 
                foreach(var colIx in colIndexes)
                {
                    var aColVal = a.GetByOriginalIndex(colIx - 1);
                    var bColVal = b.GetByOriginalIndex(colIx - 1);
                    if(aColVal.Value == null)
                    {
                        if (bColVal.Value == null) return 0;
                        return 1;
                    }
                    else if(bColVal.Value == null)
                    {
                        return -1;
                    }
                    var res = _comparer.Compare(aColVal.Value, bColVal.Value, sortOrder, context);
                    if (res != 0) return res;
                }
                return 0;
            });
            var rowIx = 0;
            foreach(var row in rows)
            {
                var cellIx = 0;
                foreach(var cell in row.ToList())
                {
                    sortedRange.SetValue(rowIx, cellIx, cell.Value);
                    cellIx++;
                }
                rowIx++;
            }
            return sortedRange;
        }

        static int CompareExcelStyle(string s1, string s2)
        {
            // Konvertera strängarna till fullständiga Unicode-koder
            var s1Codes = EnumerateRunes(s1).ToArray();
            var s2Codes = EnumerateRunes(s2).ToArray();

            // Jämför tecken för tecken baserat på deras Unicode-värden
            int minLength = Math.Min(s1Codes.Length, s2Codes.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (s1Codes[i] != s2Codes[i])
                    return s1Codes[i].CompareTo(s2Codes[i]);
            }

            // Om strängarna är identiska upp till minLength, jämför längden
            return s1Codes.Length.CompareTo(s2Codes.Length);
        }

        static IEnumerable<int> EnumerateRunes(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];

                // Kontrollera om det är en högställig surrogat (starten på ett par)
                if (char.IsHighSurrogate(c) && i + 1 < input.Length && char.IsLowSurrogate(input[i + 1]))
                {
                    // Kombinera högställig och lågställig surrogat till en fullständig Unicode-kodpunkt
                    yield return char.ConvertToUtf32(c, input[i + 1]);
                    i++; // Hoppa över nästa tecken eftersom det är en del av surrogatparet
                }
                else
                {
                    // Enskilt tecken
                    yield return c;
                }
            }
        }
    }
}
