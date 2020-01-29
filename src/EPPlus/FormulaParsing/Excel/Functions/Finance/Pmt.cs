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
    internal class Pmt : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var rate = ArgToDecimal(arguments, 0);
            var nPer = ArgToInt(arguments, 1);
            var presentValue = ArgToDecimal(arguments, 2);
            var payEndOfPeriod = false;
            var futureValue = 0d;
            if (arguments.Count() > 3) futureValue = ArgToDecimal(arguments, 3);
            if (arguments.Count() > 4) payEndOfPeriod = ArgToBool(arguments, 4);

            var result = (futureValue + presentValue * System.Math.Pow(rate + 1, nPer)) * rate
                      /
                   ((payEndOfPeriod ? rate + 1 : 1) * (1 - System.Math.Pow(rate + 1, nPer)));
       

            return CreateResult(result, DataType.Decimal);
        }

        private static double GetInterest(double rate, double remainingAmount)
        {
            return remainingAmount * rate;
        }
    }
}
