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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
	internal abstract class PivotShowAsRankBase : PivotShowAsBase
	{

		protected void CalculateRank(ExcelPivotTableDataField df, List<int> fieldIndex, PivotCalculationStore calculatedItems, bool ascending)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);
			var colFieldsStart = df.Field.PivotTable.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			var maxBfKey = 0;
			var rankItems = new SortedList<int[], object>(new ArrayComparer());
			if (record.CacheItems[df.BaseField].Count(x => x is int) > 0)
			{
				maxBfKey = (int)record.CacheItems[df.BaseField].Where(x => x is int).Max();
			}

			foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
			{
				if (IsSumBefore(key.Key, bf, fieldIndex, colFieldsStart))
				{
					calculatedItems[key.Key] = null;
				}
				else if (key.Key[keyCol] >= 0)
				{
					if (IsSumAfter(key.Key, bf, fieldIndex, colFieldsStart, 1))
					{
						rankItems.Add(key.Key, calculatedItems[key.Key]);
					}
					else
					{
						calculatedItems[key.Key] = 1D;
					}
				}
				else
				{
					calculatedItems[key.Key] = 1D;
				}
			}

			while (rankItems.Count > 0)
			{
				RankKeys(calculatedItems, rankItems, keyCol, maxBfKey, ascending);
			}
		}

		internal struct RankItem
		{
			public int[] Key { get; set; }
			public double Value { get; set; }
		}
		private void RankKeys(PivotCalculationStore calculatedItems, SortedList<int[], object> rankItems, int keyCol, int maxBfKey, bool ascending)
		{
			var startKey = (int[])rankItems.First().Key;

			var items = new List<RankItem>();
			for (int i = 0; i <= maxBfKey; i++)
			{
				var key = (int[])startKey.Clone();
				key[keyCol] = i;

				if (calculatedItems.TryGetValue(key, out object value))
				{
					if (value is double d)
					{
						items.Add(new RankItem { Key = key, Value = d });
					}
				}
			}

			double rankValue = 1;
			if (ascending)
			{
				foreach (var item in items.OrderBy(x => x.Value))
				{
					calculatedItems[item.Key] = rankValue++;
					rankItems.Remove(item.Key);
				}
			}
			else
			{
				foreach (var item in items.OrderByDescending(x => x.Value))
				{
					calculatedItems[item.Key] = rankValue++;
					rankItems.Remove(item.Key);
				}
			}
		}
	}
}