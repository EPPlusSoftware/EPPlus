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
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOf : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var pt = df.Field.PivotTable;
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            var isRowField = keyCol < pt.RowFields.Count;
            var baseLevel = isRowField ? keyCol : keyCol - pt.RowFields.Count;
            var biType = df.BaseItem == (int)ePrevNextPivotItem.Previous ? -1 : (df.BaseItem== (int)ePrevNextPivotItem.Next ? 1 : 0);
            var maxCol = pt.Fields[df.BaseField].Items.Count - 2;

            var isLowestGroupLevel = (keyCol == colStartIx - 1 || keyCol == fieldIndex.Count - 1); //If not lowest group key set value to 1 or 0 only.;

			var currentKey = GetKey(fieldIndex.Count);
            var lastIx = fieldIndex.Count-1;
            var lastItemIx = pt.Fields[fieldIndex[lastIx]].Items.Count - 1;
            while (currentKey[lastIx] < lastItemIx  || currentKey[lastIx] == PivotCalculationStore.SumLevelValue)
            {
                if (currentKey[keyCol] == PivotCalculationStore.SumLevelValue)
                {
                    showAsCalculatedItems.Add(currentKey, 0D);
                }
                else if (biType != 0 || 
                         IsSameLevelAs(currentKey, isRowField, baseLevel, keyCol, df) ||
                         currentKey[keyCol] == df.BaseItem)
                {
                    var tv = (int[])currentKey.Clone();
                    if (biType == 0)
                    {
						tv[keyCol] = df.BaseItem;
					}
                    else if (isLowestGroupLevel)
                    {
						if (biType < 0)
						{
							tv[keyCol] = tv[keyCol] == 0 ? 0 : tv[keyCol] - 1;
						}
						else if (biType > 0)
						{
							tv[keyCol] = tv[keyCol] == maxCol ? maxCol : tv[keyCol] + 1;
						}
					}

					if (calculatedItems.TryGetValue(currentKey, out object o))
                    {
                        if (o is double d)
                        {
                            if (calculatedItems.TryGetValue(tv, out object to))
                            {
                                if (to is double td)
                                {
                                    showAsCalculatedItems.Add(currentKey, d / td);
                                }
                                else
                                {
                                    showAsCalculatedItems.Add(currentKey, 0D);
                                }
                            }
                            else
                            {
                                if(isLowestGroupLevel)
                                {
                                    showAsCalculatedItems.Add(currentKey, 0D);
                                }
                                else
                                {
                                    showAsCalculatedItems.Add(currentKey, 1D);
                                }
                            }
                        }
                        else
                        {
                            if (biType == 0)
                            {
                                showAsCalculatedItems.Add(currentKey, o);
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, o);
                            }
                        }
                    }
                    else
                    {
                        if(ArrayComparer.IsEqual(currentKey, tv))
                        {                            
                            showAsCalculatedItems.Add(currentKey, 0D);
                        }
                        else
                        {
                            showAsCalculatedItems.Add(currentKey, ErrorValues.NullError);
                        }
                    }
                }
                else
                {
                    if (biType == 0)
                    {
                        showAsCalculatedItems.Add(currentKey, ErrorValues.NAError);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(currentKey, 0);
                    }
                }
                if (NextKey(ref currentKey, pt, fieldIndex) == false) break;
            }

            calculatedItems = showAsCalculatedItems;
        }
        private bool NextKey(ref int[] currentKey, ExcelPivotTable pt, List<int> fieldIndex)
        {
            currentKey = (int[])currentKey.Clone();
            int i = 0;
            currentKey[i] = (currentKey[i] == PivotCalculationStore.SumLevelValue ? 0 : currentKey[i] + 1);
			while (currentKey[i] == pt.Fields[fieldIndex[i]].Items.Count - 1)
            {
                currentKey[i] = PivotCalculationStore.SumLevelValue;
                i++;
                if (i == currentKey.Length) return false;
				currentKey[i] = (currentKey[i] == PivotCalculationStore.SumLevelValue ? 0 : currentKey[i] + 1);
			}
			return true;
        }
    }
}
