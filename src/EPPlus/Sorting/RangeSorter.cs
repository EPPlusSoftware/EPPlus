/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Sorting.Internal;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    internal class RangeSorter
    {
        public RangeSorter(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        private readonly ExcelWorksheet _worksheet;

        private void ValidateColumnArray(ExcelRangeBase range, int[] columns)
        {
            var cols = range._toCol - range._fromCol + 1;
            foreach (var c in columns)
            {
                if (c > cols - 1 || c < 0)
                {
                    throw (new ArgumentException("Cannot reference columns outside the boundries of the range. Note that column references are zero-based within the range"));
                }
            }
        }

        private void ValidateRowsArray(ExcelRangeBase range, int[] rows)
        {
            var nRows = range._toRow - range._fromRow + 1;
            foreach (var r in rows)
            {
                if (r > nRows - 1 || r < 0)
                {
                    throw (new ArgumentException("Cannot reference rows outside the boundries of the range. Note that row references are zero-based within the range"));
                }
            }
        }

        private bool[] CreateDefaultDescendingArray(int[] sortParams)
        {
            var descending = new bool[sortParams.Length];
            for (int i = 0; i < sortParams.Length; i++)
            {
                descending[i] = false;
            }
            return descending;
        }

        public void Sort(
            ExcelRangeBase range, 
            int[] columns, 
            ref bool[] descending, 
            CultureInfo culture = null, 
            CompareOptions compareOptions = CompareOptions.None, 
            Dictionary<int, string[]> customLists = null)
        {
            if (columns == null)
            {
                columns = new int[] { 0 };
            }
            ValidateColumnArray(range, columns);
            if (descending == null)
            {
                descending = CreateDefaultDescendingArray(columns);
            }
            var sortItems = SortItemFactory.Create(range);
            var comp = new EPPlusSortComparer(columns, descending, customLists, culture ?? CultureInfo.CurrentCulture, compareOptions);
            sortItems.Sort(comp);
            var wsd = new RangeWorksheetData(range);

            ApplySortedRange(range, sortItems, wsd);
        }

        public void SortLeftToRight(
            ExcelRangeBase range,
            int[] rows,
            ref bool[] descending,
            CultureInfo culture,
            CompareOptions compareOptions = CompareOptions.None,
            Dictionary<int, string[]> customLists = null
            )
        {
            if (rows == null)
            {
                rows = new int[] { 0 };
            }
            ValidateRowsArray(range, rows);
            if (descending == null)
            {
                descending = CreateDefaultDescendingArray(rows);
            }
            var sortItems = SortItemLeftToRightFactory.Create(range);
            var comp = new EPPlusSortComparerLeftToRight(rows, descending, customLists, culture ?? CultureInfo.CurrentCulture, compareOptions);
            sortItems.Sort(comp);
            var wsd = new RangeWorksheetData(range);

            ApplySortedRange(range, sortItems, wsd);
        }

        private void ApplySortedRange(ExcelRangeBase range, List<SortItem<ExcelValue>> sortItems, RangeWorksheetData wsd)
        {
            //Sort the values and styles.
            var nColumnsInRange = range._toCol - range._fromCol + 1;
            _worksheet._values.Clear(range._fromRow, range._fromCol, range._toRow - range._fromRow + 1, nColumnsInRange);
            for (var r = 0; r < sortItems.Count; r++)
            {
                for (int c = 0; c < nColumnsInRange; c++)
                {
                    var row = range._fromRow + r;
                    var col = range._fromCol + c;
                    //_worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
                    _worksheet._values.SetValue(row, col, sortItems[r].Items[c]);
                    var addr = ExcelCellBase.GetAddress(sortItems[r].Row, range._fromCol + c);
                    //Move flags
                    HandleFlags(wsd, row, col, addr);
                    //Move metadata
                    HandleMetadata(wsd, row, col, addr);

                    //Move formulas
                    HandleFormula(wsd, row, col, addr, sortItems[r].Row, col);

                    //Move hyperlinks
                    HandleHyperlink(wsd, row, col, addr);

                    //Move comments
                    HandleComment(wsd, row, col, addr);
                }
            }
        }

        private void ApplySortedRange(ExcelRangeBase range, List<SortItemLeftToRight<ExcelValue>> sortItems, RangeWorksheetData wsd)
        {
            //Sort the values and styles.
            var nRowsInRange = range._toRow - range._fromRow + 1;
            _worksheet._values.Clear(range._fromRow, range._fromCol, range._toRow - range._fromRow + 1, range._toCol);
            for (var c = 0; c < sortItems.Count; c++)
            {
                for (int r = 0; r < nRowsInRange; r++)
                {
                    var row = range._fromRow + r;
                    var col = range._fromCol + c;
                    //_worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
                    _worksheet._values.SetValue(row, col, sortItems[c].Items[r]);
                    var addr = ExcelCellBase.GetAddress(range._fromRow + r, sortItems[c].Column);
                    //Move flags
                    HandleFlags(wsd, row, col, addr);
                    //Move metadata
                    HandleMetadata(wsd, row, col, addr);

                    //Move formulas
                    HandleFormula(wsd, row, col, addr, row, sortItems[c].Column);

                    //Move hyperlinks
                    HandleHyperlink(wsd, row, col, addr);

                    //Move comments
                    HandleComment(wsd, row, col, addr);
                }
            }
        }

        private void HandleHyperlink(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Hyperlinks.ContainsKey(addr))
            {
                _worksheet._hyperLinks.SetValue(row, col, wsd.Hyperlinks[addr]);
            }
        }

        private void HandleMetadata(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Metadata.ContainsKey(addr))
            {
                _worksheet._metadataStore.SetValue(row, col, wsd.Metadata[addr]);
            }
        }

        private void HandleFlags(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Flags.ContainsKey(addr))
            {
                _worksheet._flags.SetValue(row, col, wsd.Flags[addr]);
            }
        }

        private void HandleComment(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Comments.ContainsKey(addr))
            {
                var i = wsd.Comments[addr];
                _worksheet._commentsStore.SetValue(row, col, i);
                var comment = _worksheet._comments[i];
                comment.Reference = ExcelCellBase.GetAddress(row, col);
            }
        }

        private void HandleFormula(RangeWorksheetData wsd, int row, int col, string addr, int initialRow, int initialCol)
        {
            if (wsd.Formulas.ContainsKey(addr))
            {
                _worksheet._formulas.SetValue(row, col, wsd.Formulas[addr]);
                if(wsd.Formulas[addr] is string)
                {
                    var formula = wsd.Formulas[addr].ToString();
                    var newFormula = initialRow != row ?
                        AddressUtility.ShiftAddressRowsInFormula(string.Empty, formula, initialRow, row - initialRow) :
                        AddressUtility.ShiftAddressColumnsInFormula(string.Empty, formula, initialCol, col - initialCol);
                    _worksheet._formulas.SetValue(row, col, newFormula);
                }
                else if (wsd.Formulas[addr] is int)
                {
                    int sfIx = (int)wsd.Formulas[addr];
                    var startAddr = new ExcelAddress(_worksheet._sharedFormulas[sfIx].Address);
                    var f = _worksheet._sharedFormulas[sfIx];
                    if (startAddr._fromRow > row)
                    {
                        f.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), row, f.StartCol);
                        f.StartRow = row;
                        f.Address = ExcelCellBase.GetAddress(row, startAddr._fromCol, startAddr._toRow, startAddr._toCol);
                    }
                    else if (startAddr._toRow < row)
                    {
                        f.Address = ExcelCellBase.GetAddress(startAddr._fromRow, startAddr._fromCol, row, startAddr._toCol);
                    }
                }
            }
        }

        internal void SetWorksheetSortState(ExcelRangeBase range, int[] columnsOrRows, bool[] descending, CompareOptions compareOptions, bool leftToRight, Dictionary<int, string[]> customLists)
        {
            //Set sort state
            var sortState = new SortState(_worksheet.NameSpaceManager, _worksheet);
            sortState.Ref = range.Address;
            sortState.ColumnSort = leftToRight;
            sortState.CaseSensitive = (compareOptions == CompareOptions.IgnoreCase || compareOptions == CompareOptions.OrdinalIgnoreCase);
            for (var ix = 0; ix < columnsOrRows.Length; ix++)
            {
                bool? desc = null;
                if (descending.Length > ix && descending[ix])
                {
                    desc = true;
                }
                var adr = leftToRight ?
                    ExcelCellBase.GetAddress(range._fromRow + columnsOrRows[ix], range._fromCol, range._fromRow + columnsOrRows[ix], range._toCol) :
                    ExcelCellBase.GetAddress(range._fromRow, range._fromCol + columnsOrRows[ix], range._toRow, range._fromCol + columnsOrRows[ix]);
                if (customLists != null && customLists.ContainsKey(columnsOrRows[ix]))
                {
                    sortState.SortConditions.Add(adr, desc, customLists[columnsOrRows[ix]]);
                }
                else
                {
                    sortState.SortConditions.Add(adr, desc);
                }
            }
        }
    }
}
