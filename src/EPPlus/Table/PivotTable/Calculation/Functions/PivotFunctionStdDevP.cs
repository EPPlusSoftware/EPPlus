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
    internal class PivotFunctionStdDevP : PivotFunction
    {
        internal override void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems, Dictionary<int[], int> keyCount)
        {   
            var d = GetValueDouble(value);
            if (double.IsNaN(d))
            {
                AddItemsToKeys<ExcelErrorValue>(key, dataFieldItems, keyCount, (ExcelErrorValue)value, SetError);
            }
            else
            {
                AddItemsToKeys<object>(key, dataFieldItems, keyCount, d, ValueList);
            }
        }
        internal override void Calculate(List<object> list, Dictionary<int[], object> dataFieldItems)
        {
            foreach (var key in dataFieldItems.Keys.ToArray())
            {
                if (dataFieldItems[key] is List<double> l)
                {
                    double avg = l.AverageKahan();
                    dataFieldItems[key] = Math.Sqrt(l.AverageKahan(v => Math.Pow(v - avg, 2)));
                }
            }
        }
    }
}
