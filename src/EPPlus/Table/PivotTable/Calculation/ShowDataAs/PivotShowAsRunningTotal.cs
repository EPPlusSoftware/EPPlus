using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfRowTotal : PivotShowAsBase
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
                    var rowTotal = GetRowTotal(key.Key, colStartIx, calculatedItems, out ExcelErrorValue error);
                    if (double.IsNaN(rowTotal))
                    {
                        showAsCalculatedItems.Add(key.Key,error);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key.Key, d / rowTotal);
                    }                    
                }
            }
            calculatedItems = showAsCalculatedItems;
        }
        private static double GetRowTotal(int[] key, int colStartIx, PivotCalculationStore calculatedItems, out ExcelErrorValue error)
        {
            var rowKey = (int[])key.Clone();
            for (int i = colStartIx; i < key.Length; i++)
            {
                rowKey[i] = -1;
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
