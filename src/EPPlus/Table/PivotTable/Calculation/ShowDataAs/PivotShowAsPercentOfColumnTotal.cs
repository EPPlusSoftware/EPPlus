using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfColumnTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], object> calculatedItems)
        {   
            var showAsCalculatedItems = new Dictionary<int[], object>();
            var colStartIx = df.Field.PivotTable.RowFields.Count;
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            foreach(var key in calculatedItems.Keys.ToArray())
            {
                if (calculatedItems[key] is double d)
                {
                    var colTotal = GetColumnTotal(key, colStartIx, calculatedItems, out ExcelErrorValue error);
                    if (double.IsNaN(colTotal))
                    {
                        showAsCalculatedItems.Add(key,error);
                    }
                    else
                    {
                        showAsCalculatedItems.Add(key, d / colTotal);
                    }                    
                }
            }
        }
        private static double GetColumnTotal(int[] key, int colStartIx, Dictionary<int[], object> calculatedItems, out ExcelErrorValue error)
        {
            var colKey = (int[])key.Clone();
            for (int i = 0;i < key.Length;i++)
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
