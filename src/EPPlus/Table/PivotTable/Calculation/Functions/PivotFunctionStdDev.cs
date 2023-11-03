﻿using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    //internal class StdDevItem
    //{
    //    List<double> values;
    //}
    internal class PivotFunctionStdDev : PivotFunction
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
                AddItemsToKeys<object>(key, dataFieldItems, d, ValueList);
            }
        }
        internal override void Calculate(List<object> list, Dictionary<int[], object> dataFieldItems)
        {
            foreach (var key in dataFieldItems.Keys.ToArray())
            {
                if (dataFieldItems[key] is List<double> l)
                {
                    if (l.Count > 1)
                    {
                        var avg = l.AverageKahan();
                        var sum = l.SumKahan(d => Math.Pow(d - avg, 2));
                        var div = ExcelFunction.Divide(sum, (l.Count - 1));
                        if (double.IsPositiveInfinity(div))
                        {
                            dataFieldItems[key] = ErrorValues.Div0Error;
                        }
                        else
                        {
                            dataFieldItems[key] = Math.Sqrt(div);
                        }
                    }
                    else
                    {
                        dataFieldItems[key] = ErrorValues.Div0Error;
                    }
                }
            }
        }
    }
}
