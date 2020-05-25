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
using EPPlus.PortedFunctions.LibreOffice.Finance;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal class Xirr : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(0) }, context).Select(x => (double)x);
            var dates = ArgsToDoubleEnumerable(new List<FunctionArgument> { arguments.ElementAt(1) }, context).Select(x => System.DateTime.FromOADate(x));
            var guess = 0.1;
            if(arguments.Count() > 2)
            {
                guess = ArgToDecimal(arguments, 2);
            }
            var result = XirrImpl.GetXirr(values, dates, guess);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
