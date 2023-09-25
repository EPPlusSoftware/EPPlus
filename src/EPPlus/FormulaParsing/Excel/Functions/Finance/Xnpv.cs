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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Financial,
       EPPlusVersion = "5.2",
       Description = "Calculates the net present value for a schedule of cash flows occurring at a series of supplied dates")]
    internal class Xnpv : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rate = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var arg2 = new List<FunctionArgument> { arguments.ElementAt(1) };
            var values = ArgsToDoubleEnumerable(arg2, context, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var dates = GetDates(arguments.ElementAt(2));
            if (values.Count != dates.Count())
                return CreateResult(eErrorType.Num);
            var firstDate = dates.First();
            var result = 0d;
            for(var i = 0; i < values.Count; i++)
            {
                var dt = dates.ElementAt(i);
                var val = values.ElementAt(i);
                if (dt < firstDate) return CreateResult(eErrorType.Num);
                result += val / System.Math.Pow(1d + rate, dt.Subtract(firstDate).TotalDays / 365d);
            }
            return CreateResult(result, DataType.Decimal);
        }

        private static IEnumerable<DateTime> GetDates(FunctionArgument arg)
        {
            var dates = new List<DateTime>();
            if(arg.Value is IEnumerable<FunctionArgument>)
            {
                var args = ((IEnumerable<FunctionArgument>)arg.Value).Select(x => (int)x.Value);
                foreach(var num in args)
                {
                    dates.Add(DateTime.FromOADate(num));
                }
            }
            else if (arg.Value is IRangeInfo)
            {
                foreach (var c in (IRangeInfo)arg.Value)
                {
                    var num = Convert.ToInt32(c.ValueDouble);
                    dates.Add(DateTime.FromOADate(num));
                }
            }
            return dates;
        }
    }
}
