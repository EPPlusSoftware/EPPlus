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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class RpnDefaultCompiler : RpnFunctionCompiler
    {
        public RpnDefaultCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {

        }

        internal override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            foreach (var child in children)
            {
                var compileResult = child.Compile();
                if(compileResult.DataType == DataType.ExcelRange)
                {

                }
                if (compileResult.IsResultOfSubtotal)
                {
                    var arg = new FunctionArgument(compileResult.Result, compileResult.DataType);
                    arg.SetExcelStateFlag(ExcelCellState.IsResultOfSubtotal);
                    args.Add(arg);
                }
                else
                {
                    BuildFunctionArguments(compileResult, args);     
                }
            }
            return Function.Execute(args, Context);
        }
    }
}
