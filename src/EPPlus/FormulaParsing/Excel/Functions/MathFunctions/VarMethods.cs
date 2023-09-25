/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    internal static class VarMethods
    {
        private static double Divide(double left, double right)
        {
            if (System.Math.Abs(right - 0d) < double.Epsilon)
            {
                throw new ExcelErrorValueException(eErrorType.Div0);
            }
            return left / right;
        }

        public static double Var(IEnumerable<ExcelDoubleCellValue> args)
        {
            return Var(args.Select(x => (double)x));
        }
       
        public static double Var(IEnumerable<double> args)
        {
            double avg = args.AverageKahan();
            double d = args.AggregateKahan(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
            return Divide(d, (args.Count() - 1));
        }

        public static double VarP(IEnumerable<ExcelDoubleCellValue> args)
        {
            return VarP(args.Select(x => (double)x));
        }

        public static double VarP(IEnumerable<double> args)
        {
            double avg = args.AverageKahan();
            double d = args.AggregateKahan(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
            return Divide(d, args.Count()); 
        }
    }
}
