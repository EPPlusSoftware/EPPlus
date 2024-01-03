using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionCount : PivotFunction
    {
        internal override void AddItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            if (value != null)
            {
                AddItemsToKey(key, colStartIx , dataFieldItems, keys, 1d, SumValue);
            }
        }
    }
}
