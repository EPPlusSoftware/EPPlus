/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.2
*************************************************************************************************/
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
    internal class PivotCalculationStore : IEnumerable<CacheIndexItem>
    {
        internal const int SumLevelValue = int.MaxValue;
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
        
		internal bool TryGetValue(int[] key, out object o, object emptyValue=null)
		{
            o = this[key];
            if(o==null)
            {
                o = emptyValue;
                return false;
            }
            return true;
		}

		IEnumerator<CacheIndexItem> IEnumerable<CacheIndexItem>.GetEnumerator()
		{
			return Index.GetEnumerator();
		}

		public IEnumerator GetEnumerator()
		{
            return Index.GetEnumerator();
		}

        internal void SetAllValues(ExcelErrorValue nAError)
        {
            for (int i = 0; i < Values.Count; i++)
            {
                Values[i] = ErrorValues.NAError;
            }
        }
    }
}
