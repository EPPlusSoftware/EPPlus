using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfGrandTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
        {
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            if(t is double total)
            {
                foreach(var key in calculatedItems.Index)
                {
                    if (calculatedItems[key.Key] is double d)
                    {
                        calculatedItems[key.Key] = d / total;
                    }
                }
            }
            else //Not a double, its an excel error.
            {
                foreach (var key in calculatedItems.Index)
                {
                    if (calculatedItems[key.Key] is double d)
                    {
                        calculatedItems[key.Key] = t;
                    }
                }
            }
        }
    }
}
