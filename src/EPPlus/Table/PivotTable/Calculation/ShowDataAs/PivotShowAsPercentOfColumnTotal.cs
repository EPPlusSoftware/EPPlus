using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfColumnTotal : PivotShowAsBase
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
                    var colTotal = GetColumnTotal(key.Key, colStartIx, calculatedItems, out ExcelErrorValue error);
                    if (double.IsNaN(colTotal))
                    {
                        showAsCalculatedItems.Add(key.Key,error);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key.Key, d / colTotal);
                    }                    
                }
            }
            calculatedItems = showAsCalculatedItems;
        }
        private static double GetColumnTotal(int[] key, int colStartIx, PivotCalculationStore calculatedItems, out ExcelErrorValue error)
        {
            var colKey = (int[])key.Clone();
            for (int i = 0;i < colStartIx;i++)
            {
                colKey[i] = -1;
            }
            var v = calculatedItems[colKey];
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
