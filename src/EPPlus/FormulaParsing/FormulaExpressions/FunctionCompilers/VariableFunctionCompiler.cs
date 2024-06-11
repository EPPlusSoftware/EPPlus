using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class VariableFunctionCompiler : FunctionCompiler
    {
        public VariableFunctionCompiler(ExcelFunction function) : base(function)
        {
        }

        public override CompileResult Compile(IEnumerable<CompileResult> children, ParsingContext context)
        {
            throw new NotImplementedException();
        }
    }
}
