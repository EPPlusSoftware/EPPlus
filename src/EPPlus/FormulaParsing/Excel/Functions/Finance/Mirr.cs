/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the internal rate of return for a series of periodic cash flows, considering the cost of the investment and the interest on the reinvestment of cash")]
    internal class Mirr : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(0) }, context);
            var financeRate = ArgToDecimal(arguments, 1);
            var reinvestState = ArgToDecimal(arguments, 2);
            var result = MirrImpl.MIRR(values.Select(x => (double)x).ToArray(), financeRate, reinvestState);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
