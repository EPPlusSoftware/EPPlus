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
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class RpnSumIfCompiler : RpnFunctionCompiler
    {
        internal RpnSumIfCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
            _evaluator = new ExpressionEvaluator(context);
        }

        private readonly ExpressionEvaluator _evaluator;

        internal override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            //if(children.Count() == 3 && children.ElementAt(2).HasChildren)
            //{
            //    var lastExp = children.ElementAt(2).Children.First();
            //    lastExp.IgnoreCircularReference = true;
            //    var currentAdr = Context.Scopes.Current.Address;
            //    var sumRangeAdr = new ExcelAddress(lastExp.ExpressionString);
            //    var sumRangeWs = string.IsNullOrEmpty(sumRangeAdr.WorkSheetName) ? currentAdr.WorksheetName : sumRangeAdr.WorkSheetName;
            //}
            foreach (var child in children)
            {
                var compileResult = child.Compile();
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
