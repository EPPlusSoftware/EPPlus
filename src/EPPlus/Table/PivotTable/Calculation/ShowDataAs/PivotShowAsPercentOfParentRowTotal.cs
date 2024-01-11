using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfParentRowTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
        {   
            var showAsCalculatedItems = PivotTableCalculation.GetNewCalculatedItems();
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            foreach(var key in calculatedItems.Index)
            {
                if (calculatedItems[key.Key] is double d)
                {
                    var rowTotal = GetParentRowTotal(key.Key, colStartIx, calculatedItems, out ExcelErrorValue error);
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
        private static double GetParentRowTotal(int[] key, int colStartIx, PivotCalculationStore calculatedItems, out ExcelErrorValue error)
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
