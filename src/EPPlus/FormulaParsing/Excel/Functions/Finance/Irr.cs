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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the internal rate of return for a series of periodic cash flows")]
    internal class Irr : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments[0] }, context);
            var result = default(FinanceCalcResult<double>);
            if(arguments.Count == 1)
            {
                result = IrrImpl.Irr(values.Select(x => (double)x).ToArray());
            }
            else
            {
                var guess = ArgToDecimal(arguments, 1, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                result = IrrImpl.Irr(values.Select(x => (double)x).ToArray(), guess);
            }
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
