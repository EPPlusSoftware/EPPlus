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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal class Cumprinc : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 6);
            var rate = ArgToDecimal(arguments, 0);
            var nPer = ArgToDecimal(arguments, 1);
            var presentValue = ArgToDecimal(arguments, 2);
            var startPeriod = ArgToInt(arguments, 3);
            var endPeriod = ArgToInt(arguments, 4);
            var t = ArgToInt(arguments, 5);
            if (t < 0 || t > 1) return CreateResult(eErrorType.Num);
            var func = new CumprincImpl(new PmtProvider(), new FvProvider());
            var result = func.GetCumprinc(rate, nPer, presentValue, startPeriod, endPeriod, (PmtDue)t);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }

        private static double GetInterest(double rate, double remainingAmount)
        {
            return remainingAmount * rate;
        }
    }
}
