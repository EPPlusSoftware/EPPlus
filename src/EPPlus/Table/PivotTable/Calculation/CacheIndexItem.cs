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
using System;

namespace OfficeOpenXml.Table.PivotTable.Calculation
{
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
                key+=i.ToString() + ",";
            }
            return (key.Length > 0 ? key.Substring(0, key.Length - 1):"") + " : " + Index;
		}
	}
}
