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
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class RpnLookupFunctionCompiler : RpnFunctionCompiler
    {
        internal RpnLookupFunctionCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {

        }

        internal override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            for(var x = 0; x < children.Count(); x++)
            {
                var child = children.ElementAt(x);
                //if (x > 0 || Function.SkipArgumentEvaluation)
                //{
                //    child.ParentIsLookupFunction = Function.IsLookupFuction;
                //}
                var arg = child.Compile();
                if (arg != null)
                {
                    BuildFunctionArguments(arg, arg.DataType, args);
                }
                else
                {
                    BuildFunctionArguments(null, DataType.Unknown, args);
                } 
            }
            return Function.Execute(args, Context);
        }
    }
}
