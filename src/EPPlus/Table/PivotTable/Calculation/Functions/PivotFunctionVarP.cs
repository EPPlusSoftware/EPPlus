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
    //internal class StdDevItem
    //{
    //    List<double> values;
    //}
    internal class PivotFunctionVarP : PivotFunction
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
                    double avg = l.AverageKahan();
                    double d = l.AggregateKahan(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
                    dataFieldItems[key] = ExcelFunction.Divide(d, l.Count);
                }
            }
        }
    }
}
