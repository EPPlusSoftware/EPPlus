/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.1
*************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Input;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal abstract class PivotFunction
    {
        internal abstract void AddItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys);
		internal abstract void AggregateItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys);
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
        protected static void AggregateKeys<T>(int[] key, int colStartRef, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys, T d, Action<int[], PivotCalculationStore, T> action)
        {
			int max = 1 << key.Length;
			for (int i = 1; i < max; i++)
			{
				var newKey = GetKey(key, i);
				if (IsNonTopLevel(newKey, colStartRef))
				{
					if (keys.TryGetValue(newKey, out HashSet<int[]> hs) == false)
					{
						hs = new HashSet<int[]>(new ArrayComparer());
						keys.Add(newKey, hs);
					}
					if (hs.Contains(key) == false)
					{
						hs.Add(key);
					}

				}
				action(newKey, dataFieldItems, d);
			}
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

		internal void FilterValueFields(ExcelPivotTable pivotTable, PivotCalculationStore dataFieldItems)
		{
			foreach(var valueFilter in pivotTable.Filters.Where(x=>x.Type >= ePivotTableFilterType.ValueBetween))
            {
                var keys = new List<PivotCalculationStore.CacheIndexItem>();
                foreach(var cacheItem in dataFieldItems.Index)
                {
                    var v = dataFieldItems.GetByIndex(cacheItem.Index);
					if (valueFilter.MatchNumeric(v) ==false)
                    {
                        keys.Add(cacheItem);
                    }
                }
                keys.ForEach(x => dataFieldItems.Remove(x));
            }
            
		}

		internal void Aggregate(ExcelPivotTable pivotTable, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
		{
			foreach(var key in dataFieldItems.Index.ToArray())
            {   
                AggregateItems(key.Key, pivotTable.RowFields.Count, dataFieldItems[key.Key], dataFieldItems, keys);
            }
		}
	}
}