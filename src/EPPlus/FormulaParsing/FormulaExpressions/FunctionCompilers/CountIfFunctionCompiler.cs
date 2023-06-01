using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class CountIfFunctionCompiler : FunctionCompiler
    {
        public CountIfFunctionCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
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
            if (args.Count < 2) return new CompileResult(eErrorType.Value);
            var arg2 = args.ElementAt(1);
            if((arg2.DataType == DataType.ExcelRange) && arg2.IsExcelRange)
            {
                var arg1 = args.First();
                var result = new List<object>();
                var rangeValues = arg2.ValueAsRangeInfo;
                foreach(var funcArg in rangeValues)
                {
                    var arguments = new List<FunctionArgument> { arg1 };
                    var cr = CompileResultFactory.Create(funcArg.Value);
                    BuildFunctionArguments(cr, arguments);
                    var r = Function.ExecuteInternal(arguments, Context);
                    result.Add(r.Result);
                }
                return new CompileResult(result, DataType.ExcelRange);
            }
            return Function.Execute(args, Context);
        }
    }
}
