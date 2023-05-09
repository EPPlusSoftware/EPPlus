using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class LetFunctionCompiler : FunctionCompiler
    {
        public LetFunctionCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }


        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var argIx = 0;
            var variables = new Dictionary<string, CompileResult>();
            var currentVarName = string.Empty;
            var lastChild = children.Last();
            var result = CompileResult.GetErrorResult(eErrorType.Value);
            foreach (var child in children)
            {
                if(child == lastChild)
                {
                    result = ExecuteCalculation(variables, child);
                }
                else if(argIx % 2 == 0)
                {
                    // this is a variable, variable names will arrive
                    // via Name
                    var nve = child as NamedValueExpression;
                    if (nve == null) return CompileResult.GetErrorResult(eErrorType.Value);
                    var varName = nve._unrecognizedValue;
                    if (string.IsNullOrEmpty(varName)) return CompileResult.GetErrorResult(eErrorType.Value);
                    if (variables.ContainsKey(varName)) return CompileResult.GetErrorResult(eErrorType.Value);
                    currentVarName = varName;
                }
                else
                {
                    var cr = child.Compile();
                    variables[currentVarName] = cr;
                }
                argIx++;
            }
            return result;
        }

        private CompileResult ExecuteCalculation(Dictionary<string, CompileResult> variables, Expression child)
        {
            throw new NotImplementedException();
        }
    }
}
