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
using System.Collections.Specialized;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfParentTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], HashSet<int[]>> keys, ref PivotCalculationStore calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var bf = fieldIndex.IndexOf(df.BaseField);
            foreach (var key in calculatedItems.Index)
            {
                if (calculatedItems[key.Key] is double d)
                {
                    var parentTotal = GetParentTotal(key.Key, bf, colStartIx, calculatedItems, out ExcelErrorValue error);                    
                    if (double.IsNaN(parentTotal))
                    {
                        showAsCalculatedItems.Add(key.Key, error);
                    }
                    else if(parentTotal==0)
                    {
                        showAsCalculatedItems.Add(key.Key, 0D);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key.Key, d / parentTotal);
                    }                    
                }
            }
            calculatedItems = showAsCalculatedItems;
        }
        private static double GetParentTotal(int[] key, int bf, int colStartIx, PivotCalculationStore calculatedItems, out ExcelErrorValue error)
        {
            if (bf < 0 || bf >= key.Length)
            {
                error = ErrorValues.NAError;
                return double.NaN;
            }
            if(key[bf] == PivotCalculationStore.SumLevelValue)
            {
                error = null;
                return 0D;
            }
            var start = bf+1;
            var end = start < colStartIx ? colStartIx-1 : key.Length-1;
            var parentKey = (int[])key.Clone();
            for(int i=start;i<=end;i++)
            {
                parentKey[i] = PivotCalculationStore.SumLevelValue;
            }
            var v = calculatedItems[parentKey];
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
