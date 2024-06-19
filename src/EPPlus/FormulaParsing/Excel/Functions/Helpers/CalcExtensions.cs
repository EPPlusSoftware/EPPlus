/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class CalcExtensions
    {
        public static double AggregateKahan(this IEnumerable<double> source, double seed, Func<KahanSum, double, KahanSum> func)
        {
            KahanSum sum = 0d;
            if (source == null)
            {
                throw new ArgumentNullException("source cannot be null");
            }

            if (func == null)
            {
                throw new ArgumentNullException("func cannot be null");
            }

            KahanSum result = seed;
            foreach (var element in source)
            {
                result = func(result, element);
            }

            return result;
        }

        public static double AverageKahan(this IEnumerable<double> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val;
            }
            return sum.Get() / values.Count();
        }

        public static double AverageKahan(this IList<double> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val;
            }
            return sum.Get() / values.Count;
        }

        public static double AverageKahan(this IList<double?> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val ?? 0;
            }
            return sum.Get() / values.Count(x => x.HasValue);
        }

        public static double AverageKahan(this double[] values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val;
            }
            return sum.Get() / values.Count();
        }

        public static double AverageKahan(this IEnumerable<double> values, Func<double, double> selector)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += selector.Invoke(val);
            }
            return sum.Get() / values.Count();
        }

        public static double AverageKahan(this IEnumerable<int> values, Func<int, double> selector)
        {
            KahanSum sum = 0.0;
            foreach(var val in values)
            {
                sum += selector.Invoke(val);
            }
            return sum.Get() / values.Count();
        }

        public static double SumKahan(this IEnumerable<double> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val;
            }
            return sum.Get();
        }

        public static double SumKahan(this IEnumerable<double?> values)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += val ?? 0;
            }
            return sum.Get();
        }

        public static double SumKahan(this IEnumerable<double> values, Func<double, double> selector)
        {
            KahanSum sum = 0.0;
            foreach (var val in values)
            {
                sum += selector.Invoke(val);
            }
            return sum.Get();
        }
    }
}
