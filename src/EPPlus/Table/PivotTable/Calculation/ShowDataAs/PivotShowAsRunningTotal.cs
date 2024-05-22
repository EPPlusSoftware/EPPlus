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
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsRunningTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
		{
			CalculateRunningTotal(df, fieldIndex, keys, ref calculatedItems, false);
		}

		internal static void CalculateRunningTotal(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems, bool leaveParentSum)
		{
			var bf = fieldIndex.IndexOf(df.BaseField);

            if (bf < 0 || fieldIndex.Count == 0)
			{
				calculatedItems.SetAllValues(ErrorValues.NAError);
				return;
			}
			var pt = df.Field.PivotTable;

            var colFieldsStart = pt.RowFields.Count;
			var keyCol = fieldIndex.IndexOf(df.BaseField);
			var record = pt.CacheDefinition._cacheReference.Records;
			var maxBfKey = 0;
			var isRowField = keyCol < pt.RowFields.Count;

            if (record.CacheItems[df.BaseField].Count(x => x is int) > 0)
			{
				maxBfKey = (int)record.CacheItems[df.BaseField].Where(x => x is int).Max();
			}
			var rtCalcItems = PivotTableCalculation.GetNewCalculatedItems();
			var calcTable = PivotTableCalculation.GetAsCalculatedTable(pt);
			Dictionary<int[], object> runningTotalSums = new Dictionary<int[], object>(ArrayComparer.Instance);
            for (int r = 0; r < calcTable.Count; r++)
            {
				for (int c = 0; c < calcTable[r].Count; c++)
				{
					var key = calcTable[r][c];
					if (key[keyCol]== PivotCalculationStore.SumLevelValue)
					{
						if(leaveParentSum)
						{
                            if(ExistsValueInTable(key, colFieldsStart, calculatedItems))
							{
								rtCalcItems[key] = calculatedItems[key];
							}
                        }
						else
						{
                            rtCalcItems[key] = 0D;
                        }
                        continue;
					}
					var value = 0D;
					if (calculatedItems.TryGetValue(key, out object v))
					{
						if (v is double d)
						{
							value = d;
						}
						else
						{
                            rtCalcItems[key] = v;
                        }
					}
                    var totalKey = (int[])key.Clone();
                    totalKey[keyCol] = PivotCalculationStore.SumLevelValue;
                    object calcValue;

                    if (runningTotalSums.TryGetValue(totalKey, out object tv))
					{
						if(tv is double total)
						{
                            calcValue = total + value;
                        }
						else
						{
                            calcValue = tv;
						}
                        runningTotalSums[totalKey] =  calcValue;
                    }
                    else
					{
                        calcValue = value;
                        runningTotalSums.Add(totalKey, calcValue);
                    }

                    rtCalcItems[key] = calcValue;
                }
            }
			calculatedItems = rtCalcItems;
		}
    }
}
