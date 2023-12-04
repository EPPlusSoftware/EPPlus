using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsRunningTotal : PivotShowAsBase
    {
        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref Dictionary<int[], object> calculatedItems)
        {
            var bf = fieldIndex.IndexOf(df.BaseField);
            var colFieldsStart = df.Field.PivotTable.RowFields.Count;
            ExcelErrorValue prevError = null;
            var prevValue = 0D;
            var prevKey = -1;
            foreach (var key in calculatedItems.Keys.ToArray().OrderBy(x => x, ArrayComparer.Instance))
            {
                if (IsSumBefore(key, bf, fieldIndex, colFieldsStart) || key[bf] == -1)
                {
                    calculatedItems[key] = 0D;
                }
                else
                {
                    if (prevKey != key[bf] && IsSumAfter(key, bf, fieldIndex, colFieldsStart) ==false)
                    {
                        var o = calculatedItems[key];
                        if (o is double d)
                        {
                            prevValue = d;
                            prevError = null;
                        }
                        else
                        {
                            prevValue = 0D;
                            if (o is ExcelErrorValue e)
                            {
                                prevError = e;
                            }
                            else
                            {
                                prevError = ErrorValues.ValueError;
                            }
                        }
                    }
                    else if (calculatedItems[key] is double d)
                    {
                        if (prevError == null)
                        {
                            var v = d + prevValue;
                            calculatedItems[key] = v;
                            prevValue = v;
                        }
                        else
                        {
                            calculatedItems[key] = prevError;
                        }
                    }
                }
                prevKey = key[bf];
            }
        }

        private bool IsSumBefore(int[] key, int bf, List<int> fieldIndex, int colFieldsStart)
        {
            var start = (bf >= colFieldsStart ? colFieldsStart : 0);
            for (int i = start; i <= bf; i++)
            {
                if (key[i] == -1)
                {
                    return true;
                }
            }
            return false;
        }
        private bool IsSumAfter(int[] key, int bf, List<int> fieldIndex, int colFieldsStart)
        {
            var start = (bf >= colFieldsStart ? colFieldsStart : 0);
            if (start == 0)
            {

            }
            else
            {
                for (int i = start; i <= bf; i++)
                {
                    if (key[i] == -1)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
