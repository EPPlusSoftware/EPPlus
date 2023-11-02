using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal abstract class PivotFunction
    {
        internal abstract void AddItems(int[] key, object value, Dictionary<int[], object> dataFieldItems);
        internal virtual void Calculate(List<object> list, Dictionary<int[], object> dataFieldItems) { }
        protected static bool IsNumeric(object value)
        {
            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Double:
                case TypeCode.Single:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Decimal:
                case TypeCode.DateTime:
                    return true;
                case TypeCode.Object:
                    if (value is TimeSpan ts)
                    {
                        return true;
                    }
                    return false;
                default:
                    return false;
            }
        }
        protected static double GetValueDouble(object value)
        {
            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Double:
                case TypeCode.Single:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Decimal:
                    return Convert.ToDouble(value);
                case TypeCode.DateTime:
                    return ((DateTime)value).ToOADate();
                case TypeCode.Object:
                    if (value is TimeSpan ts)
                    {
                        //return ts.TotalDays;
                        return new DateTime(ts.Ticks).ToOADate();
                    }
                    if (value is ExcelErrorValue ev)
                    {
                        return double.NaN;
                    }
                    return 0D;
                default:
                    return 0D;
            }
        }
        protected static void SumValue(int[] key, Dictionary<int[], object> dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                dataFieldItems[key] = (double)v + d;
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void MultiplyValue(int[] key, Dictionary<int[], object> dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                dataFieldItems[key] = (double)v * d;
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void AddItemsToKeys(int[] key, Dictionary<int[], object> dataFieldItems, double d, Action<int[], Dictionary<int[], object>, double> action)
        {
            action(key, dataFieldItems, d);
            for (int i = 0; i < key.Length; i++)
            {
                var newKey = (int[])key.Clone();
                newKey[i] = -1;
                action(newKey, dataFieldItems, d);
            }
            var inc = 1;
            while (inc < key.Length)
            {
                for (int i = 0; i < key.Length - inc; i++)
                {
                    var newKey = (int[])key.Clone();
                    for (int c = 0; c <= inc; c++)
                    {
                        newKey[i + c] = -1;
                    }
                    action(newKey, dataFieldItems, d);
                }
                inc++;
            }
        }
    }
}