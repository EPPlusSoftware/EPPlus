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
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class IsBlankFunctionCompiler : FunctionCompiler
    {
        public IsBlankFunctionCompiler(ExcelFunction function, ParsingContext context)
            :base(function, context)
        {
            
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            if (children.Count() != 1) return new CompileResult(eErrorType.Value);
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var firstChild = children.First();
            try
            {
                firstChild.treatEmptyAsZero = false;
                var result = firstChild.Compile(false);
                if (result.DataType == DataType.Empty)
                {
                    args.Add(new FunctionArgument(null));
                }
                else 
                {
                    args.Add(new FunctionArgument(true));
                }

            }
            catch (ExcelErrorValueException)
            {
                args.Add(new FunctionArgument(false));
            }
            return Function.Execute(args, Context);
        }
    }
}
