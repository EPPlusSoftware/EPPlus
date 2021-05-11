/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.6
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

        private static Dictionary<string, T> GetItems<T>(ExcelRangeBase r, CellStore<T> store, int fromRow, int fromCol, int toRow, int toCol)
        {
            var e = new CellStoreEnumerator<T>(store, r._fromRow, r._fromCol, r._toRow, r._toCol);
            var l = new Dictionary<string, T>();
            while (e.Next())
            {
                l.Add(e.CellAddress, e.Value);
            }
            return l;
        }

        public void Sort(ExcelRangeBase range, int[] columns, bool[] descending = null, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None, Dictionary<int, string[]> customLists = null)
        {
            if (columns == null)
            {
                columns = new int[] { 0 };
            }
            var cols = range._toCol - range._fromCol + 1;
            foreach (var c in columns)
            {
                if (c > cols - 1 || c < 0)
                {
                    throw (new ArgumentException("Can not reference columns outside the boundries of the range. Note that column reference is zero-based within the range"));
                }
            }
            var e = new CellStoreEnumerator<ExcelValue>(_worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var sortItems = new List<SortItem<ExcelValue>>();
            SortItem<ExcelValue> item = new SortItem<ExcelValue>();

            while (e.Next())
            {
                if (sortItems.Count == 0 || sortItems[sortItems.Count - 1].Row != e.Row)
                {
                    item = new SortItem<ExcelValue>() { Row = e.Row, Items = new ExcelValue[cols] };
                    sortItems.Add(item);
                }
                item.Items[e.Column - range._fromCol] = e.Value;
            }

            if (descending == null)
            {
                descending = new bool[columns.Length];
                for (int i = 0; i < columns.Length; i++)
                {
                    descending[i] = false;
                }
            }

            var comp = new EPPlusSortComparer(columns, descending, customLists, culture ?? CultureInfo.CurrentCulture, compareOptions);
            sortItems.Sort(comp);

            var flags = GetItems(range, _worksheet._flags, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var formulas = GetItems(range, _worksheet._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var hyperLinks = GetItems(range, _worksheet._hyperLinks, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var comments = GetItems(range, _worksheet._commentsStore, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var metaData = GetItems(range, _worksheet._metadataStore, range._fromRow, range._fromCol, range._toRow, range._toCol);

            //Sort the values and styles.
            _worksheet._values.Clear(range._fromRow, range._fromCol, range._toRow - range._fromRow + 1, cols);
            for (var r = 0; r < sortItems.Count; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var row = range._fromRow + r;
                    var col = range._fromCol + c;
                    //_worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
                    _worksheet._values.SetValue(row, col, sortItems[r].Items[c]);
                    var addr = ExcelCellBase.GetAddress(sortItems[r].Row, range._fromCol + c);
                    //Move flags
                    if (flags.ContainsKey(addr))
                    {
                        _worksheet._flags.SetValue(row, col, flags[addr]);
                    }
                    //Move metadata
                    if (metaData.ContainsKey(addr))
                    {
                        _worksheet._metadataStore.SetValue(row, col, metaData[addr]);
                    }

                    //Move formulas
                    if (formulas.ContainsKey(addr))
                    {
                        _worksheet._formulas.SetValue(row, col, formulas[addr]);
                        if (formulas[addr] is int)
                        {
                            int sfIx = (int)formulas[addr];
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

                    //Move hyperlinks
                    if (hyperLinks.ContainsKey(addr))
                    {
                        _worksheet._hyperLinks.SetValue(row, col, hyperLinks[addr]);
                    }

                    //Move comments
                    if (comments.ContainsKey(addr))
                    {
                        var i = comments[addr];
                        _worksheet._commentsStore.SetValue(row, col, i);
                        var comment = _worksheet._comments[i];
                        comment.Reference = ExcelCellBase.GetAddress(row, col);
                    }
                }
            }
        }

        internal void SetWorksheetSortState(ExcelRangeBase range, int[] columns, bool[] descending, CompareOptions compareOptions)
        {
            //Set sort state
            var sortState = new SortState(_worksheet.NameSpaceManager, _worksheet);
            sortState.Ref = range.Address; 
            sortState.CaseSensitive = (compareOptions == CompareOptions.IgnoreCase || compareOptions == CompareOptions.OrdinalIgnoreCase);
            for (var ix = 0; ix < columns.Length; ix++)
            {
                bool? desc = null;
                if (descending.Length > ix && descending[ix])
                {
                    desc = true;
                }
                var adr = ExcelCellBase.GetAddress(range._fromRow, range._fromCol + columns[ix], range._toRow, range._fromCol + columns[ix]);
                sortState.SortConditions.Add(adr, desc);
            }
        }
    }
}
