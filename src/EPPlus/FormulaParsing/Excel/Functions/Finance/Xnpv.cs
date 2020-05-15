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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal class Xnpv : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var rate = ArgToDecimal(arguments, 0);
            var arg2 = new List<FunctionArgument> { arguments.ElementAt(1) };
            var values = ArgsToDoubleEnumerable(arg2, context);
            var dates = GetDates(arguments.ElementAt(2), context);
            if (values.Count() != dates.Count())
                return CreateResult(eErrorType.Num);
            var firstDate = dates.First();
            var result = 0d;
            for(var i = 0; i < values.Count(); i++)
            {
                var dt = dates.ElementAt(i);
                var val = values.ElementAt(i);
                if (dt < firstDate) return CreateResult(eErrorType.Num);
                result += val / System.Math.Pow(1d + rate, dt.Subtract(firstDate).TotalDays / 365d);
            }
            return CreateResult(result, DataType.Decimal);
        }

        private IEnumerable<System.DateTime> GetDates(FunctionArgument arg, ParsingContext context)
        {
            var dates = new List<System.DateTime>();
            if(arg.Value is IEnumerable<FunctionArgument>)
            {
                var args = ((IEnumerable<FunctionArgument>)arg.Value).Select(x => (int)x.Value);
                foreach(var num in args)
                {
                    dates.Add(System.DateTime.FromOADate(num));
                }
            }
            else if (arg.Value is ExcelDataProvider.IRangeInfo)
            {
                foreach (var c in (ExcelDataProvider.IRangeInfo)arg.Value)
                {
                    var num = Convert.ToInt32(c.ValueDouble);
                    dates.Add(System.DateTime.FromOADate(num));
                }
            }
            return dates;
        }
    }
}
