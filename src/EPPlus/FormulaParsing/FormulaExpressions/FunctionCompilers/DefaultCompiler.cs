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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class DefaultCompiler : FunctionCompiler
    {
        public DefaultCompiler(ExcelFunction function)
            : base(function)
        {

        }        
        public override CompileResult Compile(IEnumerable<CompileResult> children, ParsingContext context)
        {
            var args = new List<FunctionArgument>();
            foreach (var cr in children)
            {
                BuildFunctionArguments(cr, args);
            }
            return Function.ExecuteInternal(args, context);
        }
    }
}
