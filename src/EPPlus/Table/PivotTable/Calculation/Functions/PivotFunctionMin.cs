using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionMin : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems)
        {
            double v;
            if (dataFieldItems.TryGetValue(key, out object currentValue))
            {
                if (currentValue is ExcelErrorValue) return;
                v = GetValueDouble(value);
            }
            else
            {
                v = GetValueDouble(value);
            }
            if (double.IsNaN(v))
            {
                dataFieldItems[key] = value;
            }
            else if (currentValue == null || v < (double)currentValue)
            {
                dataFieldItems[key] = v;
            }
        }
    }
}
