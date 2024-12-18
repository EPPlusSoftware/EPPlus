using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// A collection of items in a pivot table slicer.
    /// </summary>
    public class ExcelPivotTableSlicerItemCollection : IEnumerable<ExcelPivotTableSlicerItem>
    {
        //private readonly ExcelPivotTableSlicer _slicer;
        private readonly ExcelPivotTableSlicerCache _cache;
        private readonly List<ExcelPivotTableSlicerItem> _items;
        internal ExcelPivotTableSlicerItemCollection(ExcelPivotTableSlicerCache cache)
        {
            _cache = cache;
            _items = new List<ExcelPivotTableSlicerItem>();
            RefreshMe();            
        }

        /// <summary>
        /// Refresh the items from the shared items or the group items.
        /// </summary>
        public void Refresh()
        {            
            _cache._field.Cache.Refresh();
        }

        internal void RefreshMe()
        {
            var cacheItems = _cache._field.Cache.Grouping == null ? _cache._field.Cache.SharedItems : _cache._field.Cache.GroupItems;
            if (cacheItems.Count == _items.Count)
            {
                return;
            }
            else if (cacheItems.Count > _items.Count)
            {
                for (int i = _items.Count; i < cacheItems.Count; i++)
                {
                    _items.Add(new ExcelPivotTableSlicerItem(_cache, i));
                }
            }
            else
            {
                while (cacheItems.Count < _items.Count)
                {
                    _items.RemoveAt(_items.Count - 1);
                }
            }            
        }

        /// <summary>
        /// Get the enumerator for the collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelPivotTableSlicerItem> GetEnumerator()
        {
            Refresh();
            return _items.GetEnumerator();
        }

        /// <summary>
        /// Get the enumerator for the collection
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            Refresh();
            return _items.GetEnumerator();
        }
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count 
        { 
            get
            {
                return _items.Count;
            }
        }
        /// <summary>
        /// Get the value at the specific position in the collection
        /// </summary>
        /// <param name="index">The position</param>
        /// <returns></returns>
        public ExcelPivotTableSlicerItem this[int index]
        {
            get
            {
                return _items[index];
            }
        }
        /// <summary>
        /// Get the item with supplied value.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The item matching the supplied value. Returns null if no value matches.</returns>
        public ExcelPivotTableSlicerItem GetByValue(object value)
        {
            if (_cache._field.Cache._cacheLookup.TryGetValue(value??"", out int ix))
            {
                return _items[ix];
            }
            return null;
        }
        /// <summary>
        /// Get the index of the item with supplied value.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The item matching the supplied value. Returns -1 if no value matches.</returns>
        public int GetIndexByValue(object value)
        {
            if (_cache._field.Cache._cacheLookup.TryGetValue(value ?? "", out int ix))
            {
                return ix;
            }
            return -1;
        }
        /// <summary>
        /// It the object exists in the cache
        /// </summary>
        /// <param name="value">The object to check for existance</param>
        /// <returns></returns>
        public bool Contains(object value)
        {
            return _cache._field.Cache._cacheLookup.ContainsKey(value);
        }
    }
}
    