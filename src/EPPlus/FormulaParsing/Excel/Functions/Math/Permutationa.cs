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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.5",
        IntroducedInExcelVersion = "2013",
        Description = "Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects")]
    internal class Permutationa : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            number = System.Math.Floor(number);
            var numberChosen = ArgToDecimal(arguments, 1);
            if (number <= 0d || numberChosen <= 0 || number < numberChosen) return CreateResult(eErrorType.Num);
            var result = System.Math.Pow(number, numberChosen);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
