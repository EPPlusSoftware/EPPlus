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
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelPivotTableFieldItemsCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>
    {
        ExcelPivotTableField _field;
        private readonly ExcelPivotTableCacheField _cache;
        internal ExcelPivotTableFieldItemsCollection(ExcelPivotTableField field) : base()
        {
            _field = field;
            _cache = field.Cache;
        }
        /// <summary>
        /// It the object exists in the cache
        /// </summary>
        /// <param name="value">The object to check for existance</param>
        /// <returns></returns>
        public bool Contains(object value)
        {
            return _cache._cacheLookup.ContainsKey(value);
        }
        /// <summary>
        /// Get the item with the value supplied. If the value don't not exist, null is returned
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The pivot table field</returns>
        public ExcelPivotTableFieldItem GetByValue(object value)
        {
            if(_cache._cacheLookup.TryGetValue(value, out int ix))
            {
                return _list[ix];
            }
            return null;
        }
        /// <summary>
        /// Get the index of the item with the value supplied. If the value don't not exist, -1 is returned
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The index of the item</returns>
        public int GetIndexByValue(object value)
        {
            if (_cache._cacheLookup.TryGetValue(value, out int ix))
            {
                return ix;
            }
            return -1;
        }
        /// <summary>
        /// Set Hidden to false for all items in the collection
        /// </summary>
        public void ShowAll()
        {
            foreach(var item in _list)
            {
                item.Hidden = false;
            }
            _field.PageFieldSettings.SelectedItem = -1;
        }
        /// <summary>
        /// Set the ShowDetails for all items.
        /// </summary>
        /// <param name="isExpanded">The value of true is set all items to be expanded. The value of false set all items to be collapsed</param>
        public void ShowDetails(bool isExpanded=true)
        {
            if(!(_field.IsRowField || _field.IsColumnField))
            {
                //TODO: Add exception
            }
            if (_list.Count == 0) Refresh();
            foreach (var item in _list)
            {
                item.ShowDetails= isExpanded;
            }
        }
        /// <summary>
        /// Hide all items except the item at the supplied index
        /// </summary>
        public void SelectSingleItem(int index)
        {
            if(index <0 || index >= _list.Count)
            {
                throw new ArgumentOutOfRangeException("index", "Index is out of range");
            }

            foreach (var item in _list)
            {
                if (item.Type == eItemType.Data)
                {
                    item.Hidden = true;
                }
            }
            _list[index].Hidden=false;
            if(_field.IsPageField)
            {
                _field.PageFieldSettings.SelectedItem = index;
            }
        }
        /// <summary>
        /// Refreshes the data of the cache field
        /// </summary>
        public void Refresh()
        {
            _cache.Refresh();
        }
    }
    /// <summary>
    /// Base collection class for pivottable fields
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
    {
        internal List<T> _list = new List<T>();
        internal ExcelPivotTableFieldCollectionBase()
        {
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
        public virtual T this[int Index]
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
    }
}