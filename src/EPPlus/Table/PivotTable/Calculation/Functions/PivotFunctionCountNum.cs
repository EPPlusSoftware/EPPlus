using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionCountNums : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], int> keyCount)
        {
            if (IsNumeric(value))
            {
                AddItemsToKeys(key, dataFieldItems, keyCount, 1d, SumValue);
            }
        }
    }
}
