using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionSum : PivotFunction
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
                AddItemsToKeys(key, dataFieldItems, keyCount, d, SumValue);
            }
        }
        //internal override void AddItems(int[] key, int colIx, object value, PivotCalculatedItem items)
        //{
        //    double d = GetValueDouble(value);
        //    if (double.IsNaN(d))
        //    {
        //        items.SetError(key, (ExcelErrorValue)value);
        //    }
        //    else
        //    {
        //        items.Add(key, colIx, d);
        //    }
        //}

    }
}
