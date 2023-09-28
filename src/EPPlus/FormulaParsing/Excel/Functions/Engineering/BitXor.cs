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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Engineering,
        EPPlusVersion = "5.1",
        Description = "Returns a Bitwise 'Exclusive Or' of two numbers ",
        IntroducedInExcelVersion = "2013")]
    internal class BitXor : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!IsNumeric(arguments.ElementAt(0).Value) || !IsNumeric(arguments.ElementAt(1).Value)) return CreateResult(eErrorType.Value);
            if (!IsInteger(arguments.ElementAt(0).Value) || !IsInteger(arguments.ElementAt(1).Value)) return CreateResult(eErrorType.Num);
            var number1 = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var number2 = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            if (number1 < 0 || number2 < 0) return CreateResult(eErrorType.Num);
            return CreateResult(number1 ^ number2, DataType.Integer);
        }
    }
}
