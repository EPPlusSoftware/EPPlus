using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionCountNums : PivotFunction
    {
        internal override void AddItems(int[] key, int colStartIx, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            if (IsNumeric(value))
            {
                AddItemsToKey(key, colStartIx, dataFieldItems, keys, 1d, SumValue);
            }
        }
    }
}
