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
    internal class PivotShowAsPercentOfParentRowTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            foreach(var key in calculatedItems.Index)
            {
                if (calculatedItems[key.Key] is double d)
                {
                    var rowTotal = GetParentTotal(key.Key, colStartIx, calculatedItems, out ExcelErrorValue error);
                    if (double.IsNaN(rowTotal))
                    {
                        showAsCalculatedItems.Add(key.Key, error);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key.Key, d / rowTotal);
                    }                    
                }
            }

            calculatedItems = showAsCalculatedItems;
        }
        private static double GetParentTotal(int[] key, int colStartIx, PivotCalculationStore calculatedItems, out ExcelErrorValue error)
        {
            var rowKey = (int[])key.Clone();
            for(int i=colStartIx-1;i>=0;i--)
            {
                if(rowKey[i]!=PivotCalculationStore.SumLevelValue)
                {
                    rowKey[i] = PivotCalculationStore.SumLevelValue;
                    break;
                }
            }
            var v = calculatedItems[rowKey];
            if (v is ExcelErrorValue er)
            {
                error = er;
                return double.NaN;
            }
            error = null;
            return (double)v;
        }
    }
}
