﻿/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.1
*************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class PivotFunctionVar : PivotFunction
    {
        internal override void AddItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            var d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKey<ExcelErrorValue>(key, colStartIx, dataFieldItems, keys, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKey<object>(key, colStartIx, dataFieldItems, keys, d, ValueList);
            }
        }

		internal override void AggregateItems(int[] key, int colStartIx, object value, PivotCalculationStore dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
		{
			if (value is List<Double> d)
			{
				AggregateKeys<List<double>>(key, colStartIx, dataFieldItems, keys, d, DoubleListToList);
			}
			else
			{
				AggregateKeys<ExcelErrorValue>(key, colStartIx, dataFieldItems, keys, (ExcelErrorValue)value, SetError);
			}
		}

		internal override void Calculate(List<object> list, PivotCalculationStore dataFieldItems)
        {
            foreach (var key in dataFieldItems.Index.ToArray())
            {
                if (dataFieldItems[key.Key] is List<double> l)
                { 
                    if (l.Count > 1)
                    {
                        var avg = l.AverageKahan();
                        double d = l.AggregateKahan(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));

                        var div = ExcelFunction.Divide(d, (l.Count - 1));
                        if (double.IsPositiveInfinity(div))
                        {
                            dataFieldItems[key.Key] = ErrorValues.Div0Error;
                        }
                        else
                        {
                            dataFieldItems[key.Key] = div;
                        }
                    }
                    else
                    {
                        dataFieldItems[key.Key] = ErrorValues.Div0Error;
                    }
                }
            }
        }
    }
}
