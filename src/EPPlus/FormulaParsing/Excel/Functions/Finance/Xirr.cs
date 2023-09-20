/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
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
        Description = "Calculates the internal rate of return for a schedule of cash flows occurring at a series of supplied dates")]
    internal class Xirr : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(0) }, context).Select(x => (double)x);
            var dates = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(1) }, context).Select(x => DateTime.FromOADate(x));
            var guess = 0.1;
            if(arguments.Count > 2)
            {
                guess = ArgToDecimal(arguments, 2, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            }
            var result = XirrImpl.GetXirr(values, dates, guess);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
