﻿/*************************************************************************************************
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
        protected void CalculateDifferenceShared(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems, Func<double, double, object> calcFunc)
        {
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var pt = df.Field.PivotTable;
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            var isRowField = keyCol < pt.RowFields.Count;
            var baseLevel = isRowField ? keyCol : keyCol - pt.RowFields.Count;
            var biType = df.BaseItem == (int)ePrevNextPivotItem.Previous ? -1 : (df.BaseItem == (int)ePrevNextPivotItem.Next ? 1 : 0);
            var maxCol = pt.Fields[df.BaseField].Items.Count - 2;
            var existingEmptyKeys = new HashSet<int[]>(ArrayComparer.Instance);
            var isLowestGroupLevel = (keyCol == colStartIx - 1 || keyCol == fieldIndex.Count - 1); //If not lowest group key set value to 1 or 0 only.;

            var lastIx = fieldIndex.Count - 1;
            var lastItemIx = pt.Fields[fieldIndex[lastIx]].Items.Count - 1;
            var calcTable = PivotTableCalculation.GetAsCalculatedTable(pt);
            //foreach(var currentKey in pt.GetTableKeys())
            for (int r = 0; r < calcTable.Count; r++)
            {
                for (int c = 0; c < calcTable[r].Count; c++)
                {
                    var currentKey = calcTable[r][c];
                    object value = double.NaN;
                    var existsKey = calculatedItems.TryGetValue(currentKey, out value, double.NaN);
                    if (currentKey[keyCol] == PivotCalculationStore.SumLevelValue)
                    {
                        showAsCalculatedItems.Add(currentKey, 0D);
                    }
                    else if (biType != 0 ||
                             IsSameLevelAs(currentKey, isRowField, baseLevel, keyCol, df) ||
                             currentKey[keyCol] == df.BaseItem)
                    {
                        int[] relatedKey;
                        if (biType == 0)
                        {
                            if (currentKey[keyCol]==df.BaseItem)
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                                continue;
                            }
                            relatedKey = (int[])currentKey.Clone();
                            relatedKey[keyCol] = df.BaseItem;
                        }
                        else
                        {
                            if(biType<0)
                            {
                                relatedKey = GetPrevKeyFromCalculatedTable(calcTable, r, c, keyCol, isRowField);
                            }
                            else
                            {
                                relatedKey = GetNextKeyFromCalculatedTable(calcTable, r, c, keyCol, isRowField);
                            }

                            if (relatedKey == null)
                            {
                                showAsCalculatedItems.Add(currentKey, 0D);
                                continue;
                            }
                        }

                        object relatedValue = double.NaN;
                        var existsRelatedKey = calculatedItems.TryGetValue(relatedKey, out relatedValue, double.NaN);

                        if (value is double d)
                        {
                            if (relatedValue is double td)
                            {
                                showAsCalculatedItems.Add(currentKey, calcFunc(d, td));
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, ErrorValues.ValueError);
                            }
                        }
                        else
                        {
                            
                            if(relatedValue is double td)
                            {
                                showAsCalculatedItems.Add(currentKey, calcFunc(double.NaN, td));
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, ErrorValues.ValueError);
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
            }
            calculatedItems = showAsCalculatedItems;
        }

        private int[] GetPrevKeyFromCalculatedTable(List<List<int[]>> calcTable, int r, int c,int keyCol, bool isRowField)
        {
            if(isRowField)
            {
                if(r==0)
                {
                    return null;
                }
                else
                {
                    if (HasSameParent(calcTable[r][c], calcTable[r - 1][c], keyCol))
                    {
                        return calcTable[r - 1][c];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            else
            {
                if (c == 0)
                {
                    return null;
                }
                else
                {
                    if (HasSameParent(calcTable[r][c], calcTable[r][c - 1], keyCol))
                    {
                        return calcTable[r][c];
                    }
                    else
                    {
                        return null;
                    }
                }

            }
        }
        private int[] GetNextKeyFromCalculatedTable(List<List<int[]>> calcTable, int r, int c, int keyCol, bool isRowField)
        {
            if(isRowField)
            {
                if (r == calcTable.Count-1)
                {
                    return null;
                }
                else
                {
                    if (HasSameParent(calcTable[r][c],calcTable[r + 1][c], keyCol))
                    {
                        return calcTable[r + 1][c];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            else
            {
                if (c == calcTable[r].Count-1)
                {
                    return null;
                }
                else
                {
                    if (HasSameParent(calcTable[r][c], calcTable[r][c + 1], keyCol))
                    {
                        return calcTable[r][c + 1];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        private bool HasSameParent(int[] key1, int[] key2, int keyCol)
        {
            for(int i=0;i<key1.Length; i++)
            {
                if(i!=keyCol && key1[i] != key2[i])
                {
                    return false;
                }
            }
            return true;
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
        private object CalcDifference(double value, double prevValue)
        {
            return (double.IsNaN(value) ? 0D : value) - (double.IsNaN(prevValue) ? 0D : prevValue);
        }
    }
    internal class PivotShowAsDifferencePercent : PivotShowAsDifferenceBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {
            CalculateDifferenceShared(df, fieldIndex,keys, ref calculatedItems, CalcDifferencePercent);
        }
        private object CalcDifferencePercent(double value, double prevValue)
        {
            if (double.IsNaN(value) && double.IsNaN(prevValue))
            {
                return ErrorValues.NullError;
            }
            else
            {
                if(value==prevValue || double.IsNaN(value) || double.IsNaN(prevValue))
                {
                    return 0D;
                }
                else
                {
                    return (value - prevValue) / prevValue;
                }
            }
        }
    }
}
