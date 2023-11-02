﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionProduct : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems)
        {
            double d = GetValueDouble(value);
            AddItemsToKeys(key, dataFieldItems, d, MultiplyValue);
        }
    }
}