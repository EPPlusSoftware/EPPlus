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
using System.Collections;
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
<<<<<<< HEAD
        internal Dictionary<int, int> _cacheDictionary = null;
=======
        internal Lookup<int, int> _cacheLookup = null;
>>>>>>> develop7

        List<int> _hiddenItemIndex=null;
        internal ExcelPivotTableFieldItemsCollection(ExcelPivotTableField field) : base()
        {
            _field = field;            
        }
        internal void InitNewCalculation()
        {
            _hiddenItemIndex = null;
        }
        internal List<int> HiddenItemIndex
        {
            get
            {
                if (_hiddenItemIndex == null)
                {
                    _hiddenItemIndex = GetHiddenList();
                }
                return _hiddenItemIndex;
            }
        }

        private List<int> GetHiddenList()
        {
            List<int> hiddenItems = new List<int>();
            for (int i = 0; i < _list.Count; i++)
            {
                if (_list[i].Hidden)
                {
                    hiddenItems.Add(_list[i].X);
                }
            }
            return hiddenItems;
        }
        internal void InitNewCalculation()
        {
            _hiddenItemIndex = null;
        }
        internal List<int> HiddenItemIndex
        {
            get
            {
                if (_hiddenItemIndex == null)
                {
                    _hiddenItemIndex = GetHiddenList();
                }
                return _hiddenItemIndex;
            }
        }

        private List<int> GetHiddenList()
        {
            List<int> hiddenItems = new List<int>();
            for (int i = 0; i < _list.Count; i++)
            {
                if (_list[i].Hidden)
                {
                    hiddenItems.Add(_list[i].X);
                }
            }
            return hiddenItems;
        }
        /// <summary>
        /// It the object exists in the cache
        /// </summary>
        /// <param name="value">The object to check for existance</param>
        /// <returns></returns>
        public bool Contains(object value)
        {
			var cl = _field.Cache.GetCacheLookup();
			return cl.ContainsKey(value);
        }
        /// <summary>
        /// Get the item with the value supplied. If the value does not exist, null is returned.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The pivot table field</returns>
        public ExcelPivotTableFieldItem GetByValue(object value)
        {
            var cl = _field.Cache.GetCacheLookup();
            if (cl.TryGetValue(value, out int ix))
            {
<<<<<<< HEAD
                if (_cacheDictionary.TryGetValue(ix, out int i))
                {
                    return _list[i];
=======
                if (_cacheLookup.Contains(ix))
                {
                    return _list[_cacheLookup[ix].First()];
>>>>>>> develop7
                }
            }
			return null;
        }
        /// <summary>
        /// Get the index of the item with the value supplied. If the value does not exist, -1 is returned.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The index of the item</returns>
        public int GetIndexByValue(object value)
        {
<<<<<<< HEAD
			var cl = _field.Cache.GetCacheLookup();
			if (cl.TryGetValue(value, out int ix))
            {
                if(_cacheDictionary.TryGetValue(ix, out int i))
                {
                    return i;
=======
            if (value == null) return -1; 
            var cl = _field.Cache.GetCacheLookup();
			if (cl.TryGetValue(value, out int ix))
            {
                if (_cacheLookup.Contains(ix))
                {
                    return _cacheLookup[ix].First();
>>>>>>> develop7
                }
            }
            return -1;
        }
        internal void MatchValueToIndex()
        {
<<<<<<< HEAD
            var cacheLookup = _field.Cache.GetCacheLookup();
=======
            var cache = _field.Cache;
            var isGroup = cache.Grouping != null;
            var cacheLookup = cache.GetCacheLookup();
>>>>>>> develop7
            foreach (var item in _list)
            {
                var v = item.Value ?? ExcelPivotTable.PivotNullValue;
                if (item.Type == eItemType.Data && cacheLookup.TryGetValue(v, out int x))
                {
                    item.X = cacheLookup[v];
<<<<<<< HEAD
                }
=======
                }                
>>>>>>> develop7
                else
                {
                    item.X = -1;
                }
            }
<<<<<<< HEAD
            _cacheDictionary = _list.Where(x=>x.X>=0).ToDictionary(x => x.X, y => _list.IndexOf(y));
=======
            _cacheLookup = (Lookup<int,int>)_list.Where(x=> x.X >= 0).ToLookup(x => x.X, y => _list.IndexOf(y));
>>>>>>> develop7
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
            _field.Cache.Refresh();
            MatchValueToIndex();
            _hiddenItemIndex = null;
        }

		internal void Sort(eSortType sort)
		{
            var comparer = new PivotItemComparer(sort, _field);
			_list.Sort(comparer);
<<<<<<< HEAD
            _cacheDictionary = _list.Where(x=>x.X > -1).ToDictionary(x => x.X, y=>_list.IndexOf(y));
=======
            _cacheLookup = (Lookup<int,int>)_list.Where(x=>x.X > -1).ToLookup(x => x.X, y=>_list.IndexOf(y));
>>>>>>> develop7
		}

        internal ExcelPivotTableFieldItem GetByCacheIndex(int index)
        {
<<<<<<< HEAD
            if(_cacheDictionary.TryGetValue(index, out int i))
            {
                return _list[i];
            }
=======
            if (_cacheLookup.Contains(index))
            {
                return _list[_cacheLookup[index].First()];
            }

>>>>>>> develop7
            return null;
        }

        internal class PivotItemComparer : IComparer<ExcelPivotTableFieldItem>
		{
			private int _mult;
			private ExcelPivotTableField _field;
            private bool _hasGrouping;
			public PivotItemComparer(eSortType sort, ExcelPivotTableField field)
			{
				this._mult = sort==eSortType.Ascending ? 1 : -1;
				this._field = field;
                _hasGrouping = _field.Grouping != null;
			}

			public int Compare(ExcelPivotTableFieldItem x, ExcelPivotTableFieldItem y)
			{
<<<<<<< HEAD
                if (x.Type == eItemType.Data)
                {
=======
                if (x.Type == eItemType.Data && y.Type == eItemType.Data)
                {
                    if(x.Value == null) return 1;
                    if(y.Value == null) return -1;
>>>>>>> develop7
                    var xText = GetTextValue(x);
                    var yText = GetTextValue(y);
                    return xText.CompareTo(yText) * _mult;
                }
                else
                {
<<<<<<< HEAD
					return 1;
=======
					return x.Type == eItemType.Data ? -1 : 1;
>>>>>>> develop7
				}
			}

			private string GetTextValue(ExcelPivotTableFieldItem item)
			{
				if(string.IsNullOrEmpty(item.Text))
                {
					return ExcelPivotTableCacheField.GetSharedStringText(item.Value, out _);
                }
                else
                {
                    return item.Text;
                }
			}
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
        /// <summary>
        /// Returns the zero-based index of the item.
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>the zero-based index of the item in the list</returns>
        internal int IndexOf(T item)
        {
            return _list.IndexOf(item);
        }
    }
}