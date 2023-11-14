using OfficeOpenXml.FormulaParsing.Utilities;
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
            var baseKey = GetKey(fieldIndex.Count);
            double keyValue;
            var isRowField = keyCol < pt.RowFields.Count;
            var baseLevel = isRowField ? keyCol : keyCol - pt.RowFields.Count;

            ExcelErrorValue keyError = null;
            if(df.BaseItem>=0)
            {
                baseKey[keyCol] = df.BaseItem;
                var tv = calculatedItems[baseKey];
                if (tv is double d)
                {
                    keyValue = d;
                }
                else
                {
                    keyError = (ExcelErrorValue)tv;
                    keyValue = double.NaN;
                }
            }
            else
            {
                keyValue = double.NaN;
            }
            foreach(var key in calculatedItems.Keys.ToArray())
            {
                if (calculatedItems[key] is double d)
                {
                    if (df.BaseField>=0 && IsParentOf(key, isRowField, baseLevel, keyCol, df))
                    {
                        if (double.IsNaN(keyValue))
                        {
                            if (keyError != null)
                            {
                                showAsCalculatedItems.Add(key, keyError);
                            }
                            else
                            {
                                if (df.BaseField == (int)ePrevNextPivotItem.Next)
                                {
                                    //var col = key[keyCol] >= maxCol ? maxCol : key[keyCol] + 1;
                                    //if (col > maxCol) col = maxCol;

                                    showAsCalculatedItems.Add(key, d / keyValue);
                                }
                                else if (df.BaseField == (int)ePrevNextPivotItem.Previous)
                                {
                                    //showAsCalculatedItems.Add(key, d / keyValue);
                                }
                            }
                        }
                        else
                        {
                            showAsCalculatedItems.Add(key, d / keyValue);
                        }
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key, d / keyValue);
                    }
                }
                else
                {
                    showAsCalculatedItems.Add(key, ErrorValues.NullError);
                }
            }
            calculatedItems = showAsCalculatedItems;
        }

        private bool IsParentOf(int[] key, bool isRowField, int baseLevel, int keyCol, ExcelPivotTableDataField df)
        {
            if (key[keyCol] == df.BaseField) return true;
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
                    
            return false;
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
