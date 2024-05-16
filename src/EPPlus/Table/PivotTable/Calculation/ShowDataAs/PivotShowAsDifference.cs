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
    internal abstract class PivotShowAsDifferenceBase : PivotShowAsBase
    {
        protected void CalculateDifferenceShared(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems, Func<double, double, double> calcFunc)
        {
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var pt = df.Field.PivotTable;
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            var isRowField = keyCol < pt.RowFields.Count;
            var baseLevel = isRowField ? keyCol : keyCol - pt.RowFields.Count;
            var biType = df.BaseItem == (int)ePrevNextPivotItem.Previous ? -1 : (df.BaseItem == (int)ePrevNextPivotItem.Next ? 1 : 0);
            var maxCol = pt.Fields[df.BaseField].Items.Count - 2;

            var isLowestGroupLevel = (keyCol == colStartIx - 1 || keyCol == fieldIndex.Count - 1); //If not lowest group key set value to 1 or 0 only.;

            //var currentKey = GetKey(fieldIndex.Count);
            var lastIx = fieldIndex.Count - 1;
            var lastItemIx = pt.Fields[fieldIndex[lastIx]].Items.Count - 1;
           
            foreach(var currentKey in pt.GetTableKeys())
            {
                var value = calculatedItems[currentKey];
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
                        tv[keyCol]=PivotCalculationStore.SumLevelValue;
                        if (keys.TryGetValue(tv, out HashSet<int[]> keysParent))
                        {
                            if (biType < 0)
                            {
                                tv = PivotKeyUtil.GetPreviousKeyFromKeys(keysParent, keyCol, currentKey[keyCol]);

                            }
                            else if (biType > 0)
                            {
                                tv = PivotKeyUtil.GetNextKeyFromKeys(keysParent, keyCol, currentKey[keyCol]);
                            }
                            if (tv == null)
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                                continue;
                            }
                        }
                        else
                        {
                            showAsCalculatedItems.Add(currentKey, 0D);
                            continue;
                        }
                    }
                    if (value is double d)
                    {
                        if (calculatedItems.TryGetValue(tv, out object to))
                        {
                            if (to is double td)
                            {
                                showAsCalculatedItems.Add(currentKey, calcFunc(d, td));
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                            }
                        }
                        else
                        {
                            if (ExistsValueInTable(tv, colStartIx, calculatedItems))
                            {
                                showAsCalculatedItems.Add(currentKey, calcFunc(d, 0));
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                            }
                        }
                    }
                    else
                    {
                        if (calculatedItems.TryGetValue(tv, out object to))
                        {
                            if (to is double td)
                            {
                                showAsCalculatedItems.Add(currentKey, calcFunc(0, td));
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                            }
                        }
                        else
                        {
                            showAsCalculatedItems.Add(currentKey, 0D);
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
            }

            calculatedItems = showAsCalculatedItems;
        }


        private bool IsSameLevelAs(int[] key, bool isRowField, int baseLevel, int keyCol, ExcelPivotTableDataField df)
        {
            if (isRowField)
            {
                for (int i = baseLevel + 1; i < df.Field.PivotTable.RowFields.Count; i++)
                {
                    if (key[i] != PivotCalculationStore.SumLevelValue)
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                for (int i = baseLevel + 1; i < df.Field.PivotTable.ColumnFields.Count; i++)
                {
                    if (key[i] != PivotCalculationStore.SumLevelValue)
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
    internal class PivotShowAsDifference : PivotShowAsDifferenceBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            CalculateDifferenceShared(df, fieldIndex, keys, ref calculatedItems, CalcDifference);
        }
        private double CalcDifference(double value, double prevValue)
        {
            return value - prevValue;
        }
    }
    internal class PivotShowAsDifferencePercent : PivotShowAsDifferenceBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            CalculateDifferenceShared(df, fieldIndex,keys, ref calculatedItems, CalcDifferencePercent);
        }
        private double CalcDifferencePercent(double value, double prevValue)
        {
            return (value - prevValue) / prevValue;
        }
    }
}
