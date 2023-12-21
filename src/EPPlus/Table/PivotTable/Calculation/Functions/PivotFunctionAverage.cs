﻿using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionAverage : PivotFunction
    {
        internal override void AddItems(int[] key, int colStartIx, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            var d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKey<ExcelErrorValue>(key, colStartIx, dataFieldItems, keys, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKey<object>(key, colStartIx, dataFieldItems, keys, d, AverageValue);
            }
        }
		//internal override void AggregateItems(int[] key, int colStartIx, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
		//{
		//	var d = GetValueDouble(value);
		//	if (double.IsNaN(d))
		//	{
		//		AggregateKeys<ExcelErrorValue>(key, colStartIx, dataFieldItems, keys, (ExcelErrorValue)value, SetError);
		//	}
		//	else
		//	{
		//		AggregateKeys<object>(key, colStartIx, dataFieldItems, keys, d, AverageValue);
		//	}
		//}

		internal override void Calculate(List<object> list, Dictionary<int[], object> dataFieldItems)
        {
            foreach (var key in dataFieldItems.Keys.ToArray())
            {
                dataFieldItems[key] = ((AverageItem)dataFieldItems[key]).Average;
            }
        }
    }
}