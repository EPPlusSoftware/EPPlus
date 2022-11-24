using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
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
                var rangeExpr = children.ElementAt(rangeIx).Children.First();
                rangeExpr.IgnoreCircularReference = true;
                var currentAdr = Context.Scopes.Current.Address;
                //var rangeAdr = new ExcelAddress(rangeExpr.ExpressionString);
                //var rangeWs = string.IsNullOrEmpty(rangeAdr.WorkSheetName) ? currentAdr.Worksheet : rangeAdr.WorkSheetName;
                //if (currentAdr.Worksheet == rangeWs && rangeAdr.Collide(new ExcelAddress(currentAdr.Address)) != ExcelAddressBase.eAddressCollition.No)
                //{
                //    var candidateArg = children.ElementAt(rangeIx + 1)?.Children.FirstOrDefault()?.Compile().Result;
                //    if (children.ElementAt(rangeIx).HasChildren)
                //    {
                //        var functionRowIndex = (currentAdr.FromRow - rangeAdr._fromRow);
                //        var functionColIndex = (currentAdr.FromCol - rangeAdr._fromCol);
                //        var firstRangeResult = children.ElementAt(rangeIx).Children.First().Compile().Result as IRangeInfo;
                //        if (firstRangeResult != null)
                //        {
                //            var candidateRowIndex = firstRangeResult.Address._fromRow + functionRowIndex;
                //            var candidateColIndex = firstRangeResult.Address._fromCol + functionColIndex;
                //            var candidateValue = firstRangeResult.GetValue(candidateRowIndex, candidateColIndex);
                //            if (_evaluator.Evaluate(candidateArg, candidateValue?.ToString()))
                //            {
                //                if (Context.Configuration.AllowCircularReferences)
                //                {
                //                    return CompileResult.ZeroDecimal;
                //                }
                //                throw new CircularReferenceException("Circular reference detected in " + currentAdr.Address);
                //            }
                //        }

                //    }
                //}
                // todo: check circular ref for the actual cell where the SumIf formula resides (currentAdr).
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
