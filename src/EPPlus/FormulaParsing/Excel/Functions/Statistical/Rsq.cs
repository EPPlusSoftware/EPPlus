/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "6.0",
    Description = "Returns the geometric mean of an array or range of positive data.")]
    internal class Rsq : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var knownXs = ArgsToDoubleEnumerable(arguments[0], context, out ExcelErrorValue e1).ToArray();
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var knownYs = ArgsToDoubleEnumerable(arguments[1], context, out ExcelErrorValue e2).ToArray();
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var result = Math.Pow(Pearson.PearsonImpl(knownXs, knownYs), 2);
            return CreateResult(result, DataType.Decimal);
        }
    }
}