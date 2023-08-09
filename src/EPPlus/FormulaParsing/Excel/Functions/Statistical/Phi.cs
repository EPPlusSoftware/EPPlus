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
            IntroducedInExcelVersion = "2013",
            Description = "Calculates the value of the density function for a standard normal distribution for a supplied number.")]
    internal class Phi : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var n = ArgToDecimal(arguments, 0);
            var result = Math.Exp(-0.5 * Math.Pow(n, 2)) / Math.Sqrt(System.Math.PI * 2);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
