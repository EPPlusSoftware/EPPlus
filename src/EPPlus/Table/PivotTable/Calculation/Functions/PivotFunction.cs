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
        protected static void SetError(int[] key, Dictionary<int[], object> dataFieldItems, ExcelErrorValue v)
        {
            dataFieldItems[key] = v;
        }
        protected static void SumValue(int[] key, Dictionary<int[], object> dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                if(v is double cv)
                {
                    dataFieldItems[key] = cv + d;
                }
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
                if (v is double cv)
                {
                    dataFieldItems[key] = (double)v * d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void MinValue(int[] key, Dictionary<int[], object> dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object o))
            {
                if (o is double cv && d < (double)cv)
                {
                    dataFieldItems[key] = d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }

        }
        protected static void MaxValue(int[] key, Dictionary<int[], object> dataFieldItems, double d)
        {
            if (dataFieldItems.TryGetValue(key, out object o))
            {
                if (o is double cv && d > (double)cv)
                {
                    dataFieldItems[key] = d;
                }
            }
            else
            {
                dataFieldItems[key] = d;
            }
        }
        protected static void AverageValue(int[] key, Dictionary<int[], object> dataFieldItems, object value)
        {
            if (dataFieldItems.TryGetValue(key, out object v))
            {
                if (v is AverageItem ai)
                {
                    dataFieldItems[key] = ai + value;
                }
            }
            else
            {
                dataFieldItems[key] = new AverageItem((double)value);
            }
        }
        protected static void ValueList(int[] key, Dictionary<int[], object> dataFieldItems, object value)
        {
            if (dataFieldItems.TryGetValue(key, out object cv))
            {
                if (cv is List<double> l)
                {
                    l.Add((double)value);
                }
            }
            else
            {
                dataFieldItems[key] = new List<double>() { (double)value };
            }
        }
        private static void GetMinMaxValue(int[] key, Dictionary<int[], object> dataFieldItems, object value, bool isMin)
        {
            double v;
            if (dataFieldItems.TryGetValue(key, out object currentValue))
            {
                if (currentValue is ExcelErrorValue) return;
                v = GetValueDouble(value);
            }
            else
            {
                v = GetValueDouble(value);
            }
            if (double.IsNaN(v))
            {
                dataFieldItems[key] = value;
            }
            else if (isMin)
            {
                if (currentValue == null || v < (double)currentValue)
                {
                    dataFieldItems[key] = v;
                }
            }
            else
            {
                if (currentValue == null || v > (double)currentValue)
                {
                    dataFieldItems[key] = v;
                }
            }
        }

        protected static void AddItemsToKeys<T>(int[] key, Dictionary<int[], object> dataFieldItems, T d, Action<int[], Dictionary<int[], object>, T> action)
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