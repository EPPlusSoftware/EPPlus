using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation
{
    internal class PivotCalculationStore : IEnumerable
    {
        internal const int SumLevelValue = int.MaxValue;

		internal struct CacheIndexItem : IComparable<CacheIndexItem> 
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
                if (Key.Length != other.Key.Length) return Key.Length > other.Key.Length ? 1 : -1; //Key length should always be equal, but add handling for different key lengths as well.
                for (int i = 0; i < Key.Length; i++)
                {                    
                    if (Key[i] != other.Key[i])
                    {
                        return Key[i].CompareTo(other.Key[i]);
                    }
                }
                return 0;
            }
			public override string ToString()
			{
                var key = "";
                foreach(var i in Key)
                {
                    key+=i.ToString()+",";
                }
                return (key.Length > 0 ? key.Substring(0, key.Length - 1):"") + " : " + Index;
			}
		}
        internal List<object> Values { get; set; } = new List<object>();
		internal List<CacheIndexItem> Index { get; set; } = new List<CacheIndexItem>();
        public int Count 
        {
            get
            {
                return Index.Count;
            }
        }
		internal void Add (int[] key, object value)
        {
            var item=new CacheIndexItem(key);
            var ix=Index.BinarySearch(item);
            if(ix >= 0) 
            {
                throw (new ArgumentException("Key already exists"));
            }
            item.Index = Values.Count;
            Values.Add(value);
            Index.Insert(~ix, item);
        }
		internal void Add(int[] key, ExcelErrorValue errorValue)
		{
			var item = new CacheIndexItem(key);
			var ix = Index.BinarySearch(item);
			if (ix >= 0)
			{
				throw (new ArgumentException("Key already exists"));
			}
			item.Index = Values.Count;
			Values.Add(errorValue);
			Index.Insert(~ix, item);
		}
		internal object this[int[] key]
        {
            get
            {
                var item = new CacheIndexItem(key);
                var ix = Index.BinarySearch(item);
                if(ix>=0)
                {
                    return Values[Index[ix].Index];
                }
                return null;
            }
            set
            {
				var item = new CacheIndexItem(key);
				var ix = Index.BinarySearch(item);
				if (ix < 0)
				{
					Add(key, value);
				}
                else
                {
			        Values[Index[ix].Index] = value;
				}
			}
		}
        internal object GetByIndex(int index)
        {
            if(index < 0 || index >= Values.Count) 
            { 
                return null; 
            }
            var key = Index[index];
            return Values[key.Index];
        }
        internal int GetIndex(int[] key)
        {
            var item = new CacheIndexItem(key);
            return Index.BinarySearch(item);
        }
		internal bool ContainsKey(int[] key)
        {
			var item = new CacheIndexItem(key);
			return Index.BinarySearch(item) >= 0;
		}
		internal object GetPreviousValue(int[] key)
        {
            var item = new CacheIndexItem(key);
            var ix = Index.BinarySearch(item);
            if (ix >= 0)
            {
                if (ix-1 >= 0)
                {
                    return Values[Index[ix - 1].Index];
                }
                return null;
            }
            ix = ~ix - 1;
            if (ix >= 0 && ix < Values.Count)
            {
                return Values[Index[ix].Index];
            }
            return null;
        }
        internal object GetNextValue(int[] key)
        {
            var item = new CacheIndexItem(key);
            var ix = Index.BinarySearch(item);
            if (ix >= 0)
            {
                if (ix + 1 < Index.Count)
                {
                    return Values[Index[ix+1].Index];
                }
                return null;
            }
            ix = ~ix;
            if (ix >= 0 && ix < Values.Count)
            {
                return Values[Index[ix].Index];
            }
            return null;
        }
        internal void Remove(CacheIndexItem item)
        {
            Index.Remove(item);
        }
		public IEnumerator GetEnumerator()
		{
			return Index.GetEnumerator();
		}

		internal bool TryGetValue(int[] key, out object o)
		{
            o = this[key];
            return o != null;
		}
	}
}
