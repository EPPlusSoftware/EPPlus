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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal
{
    internal static class SortItemFactory
    {
        internal static List<SortItem<ExcelValue>> Create(ExcelRangeBase range)
        {
            var e = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var sortItems = new List<SortItem<ExcelValue>>();
            SortItem<ExcelValue> item = new SortItem<ExcelValue>();
            var cols = range._toCol - range._fromCol + 1;
            while (e.Next())
            {
                if (sortItems.Count == 0 || sortItems[sortItems.Count - 1].Row != e.Row)
                {
                    item = new SortItem<ExcelValue>() { Row = e.Row, Items = new ExcelValue[cols] };
                    sortItems.Add(item);
                }
                item.Items[e.Column - range._fromCol] = e.Value;
            }
            return sortItems;
        }
    }
}
