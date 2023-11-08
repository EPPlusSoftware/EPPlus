using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsPercentOfGrandTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, Dictionary<int[], object> calculatedItems)
        {
            var totalKey = GetKey(fieldIndex.Count);            
            var t = calculatedItems[totalKey];
            if(t is double total)
            {
                foreach(var key in calculatedItems.Keys.ToArray())
                {
                    if (calculatedItems[key] is double d)
                    {
                        calculatedItems[key] = d / total;
                    }
                }
            }
            else //Not a double, its an excel error.
            {
                foreach (var key in calculatedItems.Keys.ToArray())
                {
                    if (calculatedItems[key] is double d)
                    {
                        calculatedItems[key] = t;
                    }
                }
            }
        }
    }
}
