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
    /// <summary>
    /// A collection of pivot tables attached to a slicer 
    /// </summary>
    public class ExcelSlicerPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        ExcelPivotTableSlicerCache _slicerCache;
        internal ExcelSlicerPivotTableCollection(ExcelPivotTableSlicerCache slicerCache)
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
        /// <summary>
        /// The indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>The pivot table at the specified index</returns>
        public ExcelPivotTable this[int index]
        {
            get
            {
                if(index < 0 || index >= _list.Count)
                {
                    throw new IndexOutOfRangeException("Index for pivot table out of range");
                }
                return _list[index];
            }
        }
        /// <summary>
        /// Adds a new pivot table to the collection. All pivot table in this collection must share the same cache.
        /// </summary>
        /// <param name="pivotTable">The pivot table to add</param>
        public void Add(ExcelPivotTable pivotTable)
        {
            if(_list.Count > 0 && _list[0].CacheId != pivotTable.CacheId)
            {
                throw (new InvalidOperationException("Multiple Pivot tables added to a slicer must refer to the same cache."));
            }
            _list.Add(pivotTable);
            _slicerCache.UpdateItemsXml();
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
    }
}
