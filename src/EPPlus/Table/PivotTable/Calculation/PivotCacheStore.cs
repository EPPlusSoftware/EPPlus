using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation
{
    internal class PivotCacheStore 
    {
        internal struct CacheIndexItem : IComparable<CacheIndexItem> //: IEqualityComparer<CacheIndexItem>, IComparer<CacheIndexItem>
        {
            internal int[] Key { get; set; }
            internal int Index { get; set; }

            public CacheIndexItem(int[] key)
            {
                Key = key;
            }

            public int Compare(CacheIndexItem x, CacheIndexItem y)
            {
                for (int i = 0; i < x.Key.Length; i++)
                {
                    if (x.Key[i] != y.Key[i])
                    {
                        return x.Key[i].CompareTo(y.Key[i]); ;
                    }
                }
                return 0;
            }

            public bool Equals(CacheIndexItem x, CacheIndexItem y)
            {
                if (x.Key.Length != y.Key.Length) return false;
                for (int i = 0; i < x.Key.Length; i++)
                {
                    if (x.Key[i] != y.Key[i]) return false;
                }
                return true;
            }

            public int GetHashCode(CacheIndexItem obj)
            {
                int hash = 49;
                for (int i = 1; i < obj.Key.Length; i++)
                {
                    unchecked
                    {
                        hash *= 23 * Key[i].GetHashCode();
                    }
                }
                return hash;

            }

            public int CompareTo(CacheIndexItem other)
            {   
                for (int i = 0; i < Key.Length; i++)
                {
                    if (Key[i] != other.Key[i])
                    {
                        return Key[i].CompareTo(other.Key[i]);
                    }
                }
                return 0;
            }
        }
        List<object> _values=new List<object>();
        List<CacheIndexItem> _index = new List<CacheIndexItem>();

        public int Count 
        {
            get
            {
                return _values.Count;
            }
        }
        internal void Add (int[] key, object value)
        {
            var item=new CacheIndexItem(key);
            var ix=_index.BinarySearch(item);
            if(ix >= 0) 
            {
                throw (new ArgumentException("Key already exists"));
            }
            item.Index = _values.Count;
            _values.Add(value);
            _index.Insert(~ix, item);
        }
        internal object this[int[] key]
        {
            get
            {
                var item = new CacheIndexItem(key);
                var ix = _index.BinarySearch(item);
                if(ix>=0)
                {
                    return _values[_index[ix].Index];
                }
                return null;
            }
        }
        internal object GetByIndex(int index)
        {
            var key = _index[index];
            return _values[key.Index];
        }
        internal int GetIndex(int[] key)
        {
            var item = new CacheIndexItem(key);
            return _index.BinarySearch(item);
        }
        internal object GetPreviousValue(int[] key)
        {
            var item = new CacheIndexItem(key);
            var ix = _index.BinarySearch(item);
            if (ix >= 0)
            {
                if (_index[ix].Index > 0)
                {
                    return _values[_index[ix].Index - 1];
                }
            }
            ix = ~ix;
            if (ix >= 0 && ix < _values.Count)
            {
                return _values[ix];
            }
            return null;
        }
        internal object GetNextValue(int[] key)
        {
            var item = new CacheIndexItem(key);
            var ix = _index.BinarySearch(item);
            if (ix >= 0)
            {
                if (_index[ix].Index + 1 < _values.Count)
                {
                    return _values[_index[ix].Index + 1];
                }
                return null;
            }
            ix = ~ix+1;
            if (ix >= 0 && ix < _values.Count)
            {
                return _values[ix];
            }
            return null;
        }
    }
}
