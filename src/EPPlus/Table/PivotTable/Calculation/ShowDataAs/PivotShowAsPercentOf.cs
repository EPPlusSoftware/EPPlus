﻿using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercent : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref Dictionary<int[], object> calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var pt = df.Field.PivotTable;
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var keyCol = fieldIndex.IndexOf(df.BaseField);
            //var baseKey = GetKey(fieldIndex.Count);
            var isRowField = keyCol < pt.RowFields.Count;
            var baseLevel = isRowField ? keyCol : keyCol - pt.RowFields.Count;
            var biType = df.BaseItem == (int)ePrevNextPivotItem.Previous ? -1 : (df.BaseItem== (int)ePrevNextPivotItem.Next ? 1 : 0);
            var maxCol = pt.Fields[df.BaseField].Items.Count - 2;
            //if(df.BaseItem>=0 && df.BaseItem < 1048828)
            //{
            //    baseKey[keyCol] = df.BaseItem;
            //}

            var currentKey = GetKey(fieldIndex.Count);
            var lastIx = fieldIndex.Count-1;
            var lastItemIx = pt.Fields[fieldIndex[lastIx]].Items.Count - 1;
            while (currentKey[lastIx] < lastItemIx)
            {
                if (currentKey[keyCol] == -1)
                {
                    showAsCalculatedItems.Add(currentKey, 0D);
                }
                else if (IsSameLevelAs(currentKey, isRowField, baseLevel, keyCol, df) ||
                         currentKey[keyCol] == df.BaseItem)
                {
                    var tv = (int[])currentKey.Clone();
                    if (biType < 0)
                    {
                        tv[keyCol] = tv[keyCol] == 0 ? 0 : tv[keyCol] - 1;
                    }
                    else if (biType > 0)
                    {
                        tv[keyCol] = tv[keyCol] == maxCol ? maxCol : tv[keyCol] + 1;
                    }
                    else
                    {
                        tv[keyCol] = df.BaseItem;
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
                                if(biType==0)
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
                            if (biType == 0)
                            {
                                showAsCalculatedItems.Add(currentKey, ErrorValues.NullError);
                            }
                            else
                            {
                                showAsCalculatedItems.Add(currentKey, 0);
                            }
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
            while (++currentKey[i] == pt.Fields[fieldIndex[i]].Items.Count - 1)
            {
                currentKey[i] = -1;
                i++;
                if (i == currentKey.Length) return false;
            }
            return true;
        }

        private bool IsSameLevelAs(int[] key, bool isRowField, int baseLevel, int keyCol, ExcelPivotTableDataField df)
        {
            if (isRowField)
            {
                for (int i = baseLevel + 1; i < df.Field.PivotTable.RowFields.Count; i++)
                {
                    if (key[i] != -1)
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
                    if (key[i] != -1)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private int GetIndexPos(ExcelPivotTableField field)
        {
            var pt = field.PivotTable;
            if (field.IsColumnField)
            {
                for (var i = 0; i < pt.ColumnFields.Count; i++)
                {
                    if (pt.RowFields[i] == field)
                    {
                        return i;
                    }
                }
            }
            else
            {
                for (var i = 0; i < pt.RowFields.Count; i++)
                {
                    if (pt.RowFields[i] == field)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }
    }
}