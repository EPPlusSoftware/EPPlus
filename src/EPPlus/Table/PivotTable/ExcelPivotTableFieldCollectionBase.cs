/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Base collection class for pivottable fields
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
    {
        internal ExcelPivotTable _table;
        private readonly int _index;
        internal List<T> _list = new List<T>();
        internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table, int index)
        {
            _table = table;
            _index = index;
        }
        /// <summary>
        /// Gets the enumerator of the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
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
        internal void AddInternal(T field)
        {
            _list.Add(field);
        }
        internal void Clear()
        {
            _list.Clear();
        }
        /// <summary>
        /// Indexer for the  collection
        /// </summary>
        /// <param name="Index">The index</param>
        /// <returns>The pivot table field</returns>
        public T this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _list.Count)
                {
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }
                return _list[Index];
            }
        }

        public void Refresh()
        {
            var ws = _table.CacheDefinition.SourceRange.Worksheet;
            var column = _table.CacheDefinition.SourceRange._fromRow + _index;
            var toRow = _table.CacheDefinition.SourceRange._toRow;
            var hs = new HashSet
                <object>();
            //Get unique values.
            for (int row = _table.CacheDefinition.SourceRange._fromRow + 1; row <= toRow; row++)
            {
                var o = ws.GetValue(row, column);
                if (!hs.Contains(o))
                {
                    hs.Add(o);
                }
            }

            //A pivot table cashe can reference multiple Pivot tables, so we need to update them all
            foreach (var pt in _table._cacheDefinition._cacheReference._pivotTables)
            {
                var existingItems = new HashSet<string>();
                var list = pt.Fields[_index].Items._list;
                var nullItems = 0;
                for (var ix = 0; ix < list.Count; ix++)
                {
                    if (list[ix].Value != null)
                    {
                        if (!hs.Contains(list[ix].Value))
                        {
                            list.RemoveAt(ix);
                            ix--;
                        }
                        else
                        {
                            existingItems.Add(list[ix].Value.ToString());
                        }
                    }
                    else
                    {
                        nullItems++;
                    }
                }
                foreach (var c in hs)
                {
                    if (!existingItems.Contains(c.ToString()))
                    {
                        list.Insert(list.Count - nullItems, new ExcelPivotTableFieldItem() { Value = c });
                    }
                }
            }
        }
    }
}