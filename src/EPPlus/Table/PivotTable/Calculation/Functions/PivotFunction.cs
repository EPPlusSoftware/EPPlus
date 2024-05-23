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
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.Table.PivotTable.Filter;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Windows.Input;
using static OfficeOpenXml.Table.PivotTable.Calculation.PivotCalculationStore;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal abstract class PivotFunction
    {
        internal abstract void AddItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys);
		internal abstract void AggregateItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, List<bool> showTotals);
		internal virtual void Calculate(List<object> list, PivotCalculationStore dataFieldItems) 
        {
        }
        protected static bool IsNumeric(object value)
        {
            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Double:
                case TypeCode.Single:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Decimal:
                case TypeCode.DateTime:
                    return true;
                case TypeCode.Object:
                    if (value is TimeSpan ts)
                    {
                        return true;
                    }
                    return false;
                default:
                    return false;
            }
        }
        protected static double GetValueDouble(object value)
        {
            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Double:
                case TypeCode.Single:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Decimal:
                    return Convert.ToDouble(value);
                case TypeCode.DateTime:
                    return ((DateTime)value).ToOADate();
                case TypeCode.Object:
                    if (value is KahanSum ks)
                    {
                        return ks;
                    }
                    if (value is TimeSpan ts)
                    {
                        //return ts.TotalDays;
                        return new DateTime(ts.Ticks).ToOADate();
                    }
                    if (value is ExcelErrorValue ev)
                    {
                        return double.NaN;
                    }
                    return 0D;
                default:
                    return 0D;
            }
        }
        protected static void SetError(int[] key, PivotCalculationStore dataFieldItems, ExcelErrorValue v)
        {
            dataFieldItems[key] = v;
        }
        protected static void SumValue(int[] key, PivotCalculationStore dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                if(v is KahanSum cv)
                {
                    dataFieldItems[key] = cv + d;
                }
            }
            else
            {
                dataFieldItems[key] = new KahanSum(d);
            }
        }
		protected static void CountValue(int[] key, PivotCalculationStore dataFieldItems, double c)
		{
			if (dataFieldItems.TryGetValue(key, out object v))
			{
				if (v is double cv)
				{
					dataFieldItems[key] = cv + c;
				}
			}
			else
			{
				dataFieldItems[key] = c;
			}
		}
		protected static void MultiplyValue(int[] key, PivotCalculationStore dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                if (v is double cv)
                {
                    dataFieldItems[key] = (double)v * d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void MinValue(int[] key, PivotCalculationStore dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object o))
            {
                if (o is double cv && d < (double)cv)
                {
                    dataFieldItems[key] = d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }

        }
        protected static void MaxValue(int[] key, PivotCalculationStore dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object o))
            {
                if (o is double cv && d > (double)cv)
                {
                    dataFieldItems[key] = d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void AverageValue(int[] key, PivotCalculationStore dataFieldItems, AverageItem value)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                if (v is AverageItem ai)
                {
                    dataFieldItems[key] = ai + value;
                }
            }
            else
            {
                dataFieldItems[key] = value;
            }
        }
		protected static void AverageValue(int[] key, PivotCalculationStore dataFieldItems, double value)
		{
			if (dataFieldItems.TryGetValue(key, out object v))
			{
				if (v is AverageItem ai)
				{
					dataFieldItems[key] = ai + value;
				}
			}
			else
			{
				dataFieldItems[key] = new AverageItem(value);
			}
		}
		protected static void ValueList(int[] key, PivotCalculationStore dataFieldItems, object value)
        {
            if (dataFieldItems.TryGetValue(key, out object cv))
            {
                if (cv is List<double> l)
                {
                    l.Add((double)value);
                }
            }
            else
            {
                dataFieldItems[key] = new List<double>() { (double)value };
            }
        }
		protected static void DoubleListToList(int[] key, PivotCalculationStore dataFieldItems, List<double> list)
		{
			if (dataFieldItems.TryGetValue(key, out object cv))
			{
				if (cv is List<double> l)
				{
					l.AddRange(list);
				}
			}
			else
			{
                dataFieldItems[key] = new List<double>(list);
			}
		}
		private static void GetMinMaxValue(int[] key, PivotCalculationStore dataFieldItems, object value, bool isMin)
        {
            double v;
            if (dataFieldItems.TryGetValue(key, out object currentValue))
            {
                if (currentValue is ExcelErrorValue) return;
                v = GetValueDouble(value);
            }
            else
            {
                v = GetValueDouble(value);
            }

            if (double.IsNaN(v))
            {
                dataFieldItems[key] = value;
            }
            else if (isMin)
            {
                if (currentValue == null || v < (double)currentValue)
                {
                    dataFieldItems[key] = v;
                }
            }
            else
            {
                if (currentValue == null || v > (double)currentValue)
                {
                    dataFieldItems[key] = v;
                }
            }
        }
        protected static void AddItemsToKey<T>(int[] key, int colStartRef, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, T d, Action<int[], PivotCalculationStore, T> action)
        {
            if (key.Length == 0)
            {
                HashSet<int[]> hs;
                if (keys.Count == 0)
                {
                    hs = new HashSet<int[]>(new ArrayComparer());
                    keys.Add(key, hs);
                }
                else
                {
                    hs = keys[key];
                }
                hs.Add(key);
                action(key, dataFieldItems, d);
                return;
            }
            bool newUniqeKey = dataFieldItems.ContainsKey(key)==false;
            action(key, dataFieldItems, d);
        }
        protected static void AggregateKeys<T>(int[] key, int colStartRef, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, T d, Action<int[], PivotCalculationStore, T> action, List<bool> showTotals)
        {
            //TODO: Check if field should be aggregated

            if (colStartRef > 0 && colStartRef < key.Length)
            {
                AddAggregatedKey(key, key.Length-1, colStartRef, dataFieldItems, keys, d, action, key);
                int[] newKey = (int[])key.Clone();
                for (int c = colStartRef-1; c >= 0; c--)
                {
                    newKey[c] = PivotCalculationStore.SumLevelValue;
                    action(newKey, dataFieldItems, d);
                    AddToKeys(keys, newKey, key);
                    AddAggregatedKey(newKey, key.Length-1, colStartRef, dataFieldItems, keys, d, action, key);
                    newKey = (int[])newKey.Clone();
                }
            }
            else
            {
                AddAggregatedKey(key, key.Length-1, 0, dataFieldItems, keys, d, action, key);
            }
        }

        private static void AddAggregatedKey<T>(int[] key, int highIx,int lowIx, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, T d, Action<int[], PivotCalculationStore, T> action, int[] uniqueKey)
        {
            int[] newKey = (int[])key.Clone();
            for (int c = highIx; c >= lowIx; c--)
            {
                newKey[c] = PivotCalculationStore.SumLevelValue;
                action(newKey, dataFieldItems, d);
                AddToKeys(keys, newKey, uniqueKey);
                newKey = (int[])newKey.Clone();
            }
        }

        private static void AddToKeys(Dictionary<int[], HashSet<int[]>> keys, int[] sumKey, int[] key)
        {
            if (keys.TryGetValue(sumKey, out HashSet<int[]> hs) == false)
            {
                hs = new HashSet<int[]>(new ArrayComparer());
                keys.Add(sumKey, hs);
            }
            if (hs.Contains(key) == false)
            {
                hs.Add(key);
            }
        }

        private static void AddKey<T>(int[] key, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, T d, Action<int[], PivotCalculationStore, T> action, int[] newKey)
        {
        }

        internal static bool IsNonTopLevel(int[] newKey, int colStartRef)
        {
           if(colStartRef > 0 && newKey[0] == PivotCalculationStore.SumLevelValue && HasSumLevel(newKey, 1, colStartRef)==false)
            {
                return true;
            }
            if (colStartRef < newKey.Length && newKey[colStartRef] == PivotCalculationStore.SumLevelValue && HasSumLevel(newKey, colStartRef+1, newKey.Length) == false)
            {
                return true;
            }
            return false;
        }
        private static bool HasSumLevel(int[] newKey, int start, int end)
        {
            for(int i = start; i < end; i++)
            {
                if (newKey[i] != PivotCalculationStore.SumLevelValue) return false;
            }
            return true;
        }

        private static int[] GetGrandTotalKey(int size)
        {
            var newKey = new int[size];
            for (int i = 0; i < size; i++)
            {
                newKey[i] = PivotCalculationStore.SumLevelValue;
            }
            return newKey;
        }

        private static int[] GetKey(int[] key, int pos)
        {
            var newKey = (int[])key.Clone();
            for (int i = 0; i < key.Length; i++)
            {
                if (((1 << i) & pos) != 0)
                {
                    newKey[i] = PivotCalculationStore.SumLevelValue;
                }
            }
            return newKey;
        }
        private static bool ShouldAggregateKey(int[] key, int colStartRef, int pos, List<bool> showTotals)
        {
            for (int i = 0; i < key.Length; i++)
            {
                if (showTotals[i] == false && ((1 << i) & pos) == 0)
                {
                    return false;
                }
            }
            return true;
        }
        internal void FilterValueFields(ExcelPivotTable pivotTable, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, List<int> fieldIndex)
        {
            foreach (var valueFilter in pivotTable.Filters.Where(x => x.Type >= ePivotTableFilterType.ValueBetween))
            {
                var keyIx = fieldIndex.IndexOf(valueFilter.Fld);
                var startIx = keyIx < pivotTable.RowFields.Count ? 0 : pivotTable.RowFields.Count;
                var keySize = keyIx - startIx + 1;
                var filterItems = new PivotCalculationStore();
                foreach (CacheIndexItem cacheItem in dataFieldItems.Index)
                {
                    var newKey = new int[keySize];
                    for (int i = startIx; i <= keyIx; i++)
                    {
                        newKey[i] = cacheItem.Key[startIx + i];
                    }
                    AddItems(newKey, pivotTable.RowFields.Count, dataFieldItems.Values[cacheItem.Index], filterItems, keys);
                }

                var keysToRemove = new List<int[]>();
                if(valueFilter.Type == ePivotTableFilterType.Sum ||
                   valueFilter.Type == ePivotTableFilterType.Count ||
                   valueFilter.Type == ePivotTableFilterType.Percent)
                {
                    var totDict = GetTop10TotalDictionary(filterItems);
                    var totSum = GetTop10SumDict(totDict, valueFilter);
                    foreach (CacheIndexItem item in filterItems.OrderBy(x => x.Key, new ArrayComparer()))
                    {
                        var pk = GetParentKey(item.Key);
                        HandleTopBottom(valueFilter, filterItems, keysToRemove, totSum, item, pk);
                    }
                }
                else
                {
                    foreach (CacheIndexItem item in filterItems)
                    {
                        if (valueFilter.MatchNumeric(filterItems[item.Key]) == false)
                        {
                            keysToRemove.Add(item.Key);
                        }
                    }
                }

                for (int i = 0; i < dataFieldItems.Index.Count; i++)
                {
                    if (SumKeyMatches(dataFieldItems.Index[i].Key, startIx, keyIx, keysToRemove))
                    {
                        dataFieldItems.Index.RemoveAt(i--);
                    }
                }
            }
        }
		private Dictionary<int[], double> GetTop10SumDict(Dictionary<int[], List<double>> totDict, ExcelPivotTableFilter valueFilter)
		{
			var dic=new Dictionary<int[], double>(ArrayComparer.Instance);
            foreach (var key in totDict.Keys)
            {
                var l = totDict[key];
                var isTop = ((ExcelTop10FilterColumn)valueFilter.Filter).Top;
                switch (valueFilter.Type)
                {
                    case ePivotTableFilterType.Count:
                        var v = Convert.ToInt32(valueFilter.Value1)-1;
                        if(l.Count >= v)
                        {
                            if(isTop)
                            {
								dic.Add(key, l.OrderByDescending(x => x).ElementAt(v));
							}
							else
                            {
								dic.Add(key, l.OrderBy(x => x).ElementAt(v));
							}
                        }
                        else
                        {
                            dic.Add(key, isTop?double.MaxValue : double.MinValue);
                        }
                        break;
					case ePivotTableFilterType.Sum:
						var d = Convert.ToDouble(valueFilter.Value1) - 1;
						AddBreakItem(dic, key, l, isTop, d);
						break;
					case ePivotTableFilterType.Percent:
						var p = Convert.ToDouble(valueFilter.Value1);
						var sum = l.Sum();
						AddBreakItem(dic, key, l, isTop, sum * (p / 100));

						break;
                }
            }
            return dic;
		}

		private static void AddBreakItem(Dictionary<int[], double> dic, int[] key, List<double> l, bool isTop, double d)
		{
			var sum = new KahanSum(0d);
			foreach (var sv in (isTop ? l.OrderByDescending(x => x) : l.OrderBy(x => x)))
			{
				sum += sv;
				if (sum >= d)
				{
					dic.Add(key, sv);
					break;
				}
			}
		}

		private static void HandleTopBottom(ExcelPivotTableFilter valueFilter, PivotCalculationStore filterItems, List<int[]> keysToRemove, Dictionary<int[], double> totSum, CacheIndexItem item, int[] pk)
		{
			if (filterItems[item.Key] is KahanSum d)
			{
				var sum = totSum[pk];
                var isTop = ((ExcelTop10FilterColumn)valueFilter.Filter).Top;
				if (isTop==false && d.Get() > sum)
				{
					keysToRemove.Add(item.Key);
				}
				else if (isTop && d.Get() < sum)
				{
					keysToRemove.Add(item.Key);
				}
			}
			else
			{
				keysToRemove.Add(item.Key);
			}
		}

		private Dictionary<int[], List<double>> GetTop10TotalDictionary(PivotCalculationStore filterItems)
		{
            var di=new Dictionary<int[], List<double>>(ArrayComparer.Instance);
			foreach (CacheIndexItem fi in filterItems)
            {
				var parentKey = GetParentKey(fi.Key);
                if(!di.TryGetValue(parentKey, out var l) ) 
                {
                    l=new List<double>();
                    di.Add(parentKey, l);
                }

                if (filterItems[fi.Key] is KahanSum d)
                {
                    l.Add(d.Get());
                }
			}
            return di;
		}

		private int[] GetParentKey(int[] key)
		{
			if(key.Length <=1)
            {
                return new int[0];
            }
            else
            {
                var newKey = new int[key.Length-1];
                for(int i=1;i<key.Length; i++)
                {
                    newKey[i] = key[i];
                }
                return newKey;
            }
		}

		private bool SumKeyMatches(int[] key, int startIx, int keyIx, List<int[]> keysToRemove)
		{
			foreach(var rk in keysToRemove)
            {
                var match = true;
                for(int i=startIx;i<=keyIx;i++)
                {
                    if (rk[i - startIx] != key[i])
                    {
                        match = false; 
                    }
                }
                if (match) return true;
            }
            return false;
		}
		private void RemoveSumLevels(ref PivotCalculationStore dataFieldItems)
		{
            if (dataFieldItems.Index.Count == 0) return;
            
            var keyLen = dataFieldItems.Index[0].Key.Length;

			for (int i=0;i<dataFieldItems.Index.Count; i++)
            {
                if (HasSumLevel(dataFieldItems.Index[i].Key,0, keyLen))
                {
                    dataFieldItems.Index.RemoveAt(i--);
                }
            }
		}

		private bool ParentMatch(List<int[]> matchingSumLevels, int[] key)
		{
			foreach(var pKey in matchingSumLevels)
            {
                if(KeyMatches(key, pKey)==false)
                {
                    return true;
                }
            }
            return false;
		}

		private bool KeyMatches(int[] key, int[] pKey)
		{
			for(int i=0; i<key.Length; i++)
            {
                if(key[i] != pKey[i] && pKey[i] < PivotCalculationStore.SumLevelValue)
                {
                    return false;
                }
            }
			return true;
		}

		internal void Aggregate(ExcelPivotTable pivotTable, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
		{
            var showTotals = pivotTable.RowFields.Union(pivotTable.ColumnFields).Select(x => x.SubTotalFunctions != eSubTotalFunctions.None).ToList();
            foreach(var key in dataFieldItems.Index.ToArray())
            {   
                AggregateItems(key.Key, pivotTable.RowFields.Count, dataFieldItems[key.Key], dataFieldItems, keys, showTotals);
            }
		}
	}
}