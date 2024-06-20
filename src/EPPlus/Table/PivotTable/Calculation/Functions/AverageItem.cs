/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.2
*************************************************************************************************/using OfficeOpenXml.FormulaParsing.Excel.Operators;
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
