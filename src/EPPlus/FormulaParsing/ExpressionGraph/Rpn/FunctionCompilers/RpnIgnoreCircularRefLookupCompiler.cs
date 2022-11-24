using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class RpnIgnoreCircularRefLookupCompiler : RpnLookupFunctionCompiler
    {
        internal RpnIgnoreCircularRefLookupCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }

        internal override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            //foreach(var child in children)
            //{
            //    child.IgnoreCircularReference = true;
            //}
            return base.Compile(children);
        }
    }
}
