using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionProduct : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], int> keyCount)
        {
            double d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKeys<ExcelErrorValue>(key, dataFieldItems, keyCount, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKeys(key, dataFieldItems, keyCount, d, MultiplyValue);
            }
        }
    }
}
