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
    internal class Ddb : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            var cost = ArgToDecimal(arguments, 0);
            var salvage = ArgToDecimal(arguments, 1);
            var life = ArgToDecimal(arguments, 2);
            var period = ArgToDecimal(arguments, 3);
            var factor = 2d;
            if(arguments.Count() >= 5)
            {
                factor = ArgToDecimal(arguments, 4);
            }

            if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0)
                return CreateResult(eErrorType.Num);

            var result = DdbImpl.Ddb(cost, salvage, life, period, factor);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
