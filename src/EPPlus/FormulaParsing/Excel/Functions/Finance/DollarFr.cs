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
        EPPlusVersion = "5.5",
        Description = "Converts a dollar price expressed as a decimal, into a dollar price expressed as a fraction")]
    internal class DollarFr : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var decimalDollar = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var fractionDec = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if(e2 != null) return CreateResult(e2.Type);
            var fraction = Math.Floor(fractionDec);
            if (fraction < 0d) return CompileResult.GetErrorResult(eErrorType.Num);
            if (fraction == 0d) return CompileResult.GetErrorResult(eErrorType.Div0);
            var result = Math.Floor(decimalDollar);
            result += (decimalDollar % 1) * System.Math.Pow(10, -System.Math.Ceiling(System.Math.Log(fraction) / System.Math.Log(10))) * fraction;
            return CreateResult(result, DataType.Decimal);
        }
    }
}
