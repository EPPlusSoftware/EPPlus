using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal struct AverageItem
    {
        public AverageItem(double sum)
        {
            Sum = sum;
            Count = 1;
        }
        public KahanSum Sum { get; set; }
        public int Count { get; set; }
        public double Average 
        { 
            get 
            {               
                return Sum.Get() / Count;
            } 
        }
        public static AverageItem operator + (AverageItem a1, AverageItem a2)
        {
            a1.Sum += a2.Sum;
            a1.Count += a2.Count;
            return a1;
        }
        public static AverageItem operator +(AverageItem a1, double value)
        {
            a1.Sum += value;
            a1.Count++;
            return a1;
        }
        public static AverageItem operator +(AverageItem a1, object value)
        {
            a1.Sum += Convert.ToDouble(value);
            a1.Count++;
            return a1;
        }
    }
}
