using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;

namespace OfficeOpenXml.Table.PivotTable.Calculation.Functions
{
    internal class AverageItem
    {
        public AverageItem(double sum) 
        {
            Sum = sum;
            Count = 1;
        }

		public double Sum { get; set; }
		public int Count { get; set; }
        public double Average 
        { 
            get 
            {               
                return Sum / Count;
            } 
        }
        public static AverageItem operator + (AverageItem a1, AverageItem a2)
        {
            return new AverageItem(a1.Sum + a2.Sum) { Count = a1.Count + a2.Count };
        }
        public static AverageItem operator +(AverageItem a1, double value)
        {
			return new AverageItem(a1.Sum + value) { Count = a1.Count + 1 };
		}
		public static AverageItem operator +(AverageItem a1, object value)
        {
			return new AverageItem(a1.Sum + Convert.ToDouble(value)) { Count = a1.Count + 1 };
        }
    }
}
