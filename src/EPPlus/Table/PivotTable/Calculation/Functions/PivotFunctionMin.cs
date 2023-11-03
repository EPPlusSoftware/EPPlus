﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionMin : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems)
        {
            var d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKeys<ExcelErrorValue>(key, dataFieldItems, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKeys(key, dataFieldItems, d, MinValue);
            }
        }
    }
}
