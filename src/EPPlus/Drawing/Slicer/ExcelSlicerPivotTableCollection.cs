/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelSlicerPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        ExcelPivotTableSlicerCache _slicerCache;
        public ExcelSlicerPivotTableCollection(ExcelPivotTableSlicerCache slicerCache)
        {
            _slicerCache = slicerCache;
        }
        internal List<ExcelPivotTable> _list=new List<ExcelPivotTable>();
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public ExcelPivotTable this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        public void Add(ExcelPivotTable table)
        {
            if(_list.Count > 0 && _list[0].CacheId != table.CacheId)
            {
                throw (new InvalidOperationException("Multiple Pivot tables added to a slicer must refer to the same cache."));
            }
            _list.Add(table);
            _slicerCache.UpdateItemsXml();
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
    }
}
