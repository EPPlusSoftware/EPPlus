/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/14/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "7.2",
        Description = "Assigns names to calculation results, allowing storing intermediate calculations, values, or defining names inside a formula",
        IntroducedInExcelVersion = "Office365")]
    internal class LetFunction : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override bool ReturnsReference => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // this function is precalculated in the calc engine and we just need to pick up the result of the last arg.
            var result = CompileResultFactory.Create(arguments.Last().Value);
            return result;
        }

        public override string NamespacePrefix
        {
            get
            {
                return "_xlfn.";
            }
        }
    }
}
