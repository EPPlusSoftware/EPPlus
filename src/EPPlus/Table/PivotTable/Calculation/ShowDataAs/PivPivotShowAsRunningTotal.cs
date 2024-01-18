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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsRunningTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
		{
			CalculateRunningTotal(df, fieldIndex, ref calculatedItems, false);
		}

		internal static void CalculateRunningTotal(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems, bool leaveParentSum)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);
			var colFieldsStart = df.Field.PivotTable.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			var maxBfKey = 0;
			if (record.CacheItems[df.BaseField].Count(x => x is int) > 0)
			{
				maxBfKey = (int)record.CacheItems[df.BaseField].Where(x => x is int).Max();
			}

			foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
			{
				if (IsSumBefore(key.Key, bf, fieldIndex, colFieldsStart))
				{
					if(!(leaveParentSum == true && key.Key[keyCol] == PivotCalculationStore.SumLevelValue))
					{
						calculatedItems[key.Key] = null;
					}
				}
				else if (IsSumAfter(key.Key, bf, fieldIndex, colFieldsStart) == true)
				{
					if (key.Key[keyCol] > 0)
					{
						var prevKey = GetPrevKey(key.Key, keyCol);
						if (calculatedItems.ContainsKey(prevKey))
						{
							if (calculatedItems[key.Key] is double current)
							{
								if (calculatedItems[prevKey] is double prev)
								{
									calculatedItems[key.Key] = current + prev;
								}
								else
								{
									calculatedItems[key.Key] = calculatedItems[prevKey]; //The prev key is an error, set the value to that error.
								}
							}
						}
					}

					if (key.Key[keyCol] < maxBfKey)
					{
						var nextKey = GetNextKey(key.Key, keyCol);
						while (nextKey[keyCol] < maxBfKey && calculatedItems.ContainsKey(nextKey) == false)
						{
							calculatedItems[nextKey] = calculatedItems[key.Key];
							nextKey = GetNextKey(nextKey, keyCol);
						}
					}
				}
			}
		}
	}
}
