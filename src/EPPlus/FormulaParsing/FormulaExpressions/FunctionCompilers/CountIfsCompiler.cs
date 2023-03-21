using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class CountIfsCompiler : FunctionCompiler
    {
        public CountIfsCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
            _evaluator =  new ExpressionEvaluator(context);
        }

        private readonly ExpressionEvaluator _evaluator;

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            for(var rangeIx = 0; rangeIx < children.Count(); rangeIx += 2)
            {
                //var rangeExpr = children.ElementAt(rangeIx).Children.First();
                //rangeExpr.IgnoreCircularReference = true;
                //var currentAdr = Context.Scopes.Current.Address;
            }
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
