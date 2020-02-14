using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class IgnoreCircularRefLookupCompiler : LookupFunctionCompiler
    {
        public IgnoreCircularRefLookupCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            foreach(var child in children)
            {
                child.IgnoreCircularReference = true;
            }
            return base.Compile(children);
        }
    }
}
