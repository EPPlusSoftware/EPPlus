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
        internal override void AddItems(int[] key, int colStartIx, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], HashSet<int[]>> keys)
        {
            var d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKeys<ExcelErrorValue>(key, colStartIx, dataFieldItems, keys, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKeys<object>(key, colStartIx, dataFieldItems, keys, d, ValueList);
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
                        double d = l.AggregateKahan(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));

                        var div = ExcelFunction.Divide(d, (l.Count - 1));
                        if (double.IsPositiveInfinity(div))
                        {
                            dataFieldItems[key] = ErrorValues.Div0Error;
                        }
                        else
                        {
                            dataFieldItems[key] = div;
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
