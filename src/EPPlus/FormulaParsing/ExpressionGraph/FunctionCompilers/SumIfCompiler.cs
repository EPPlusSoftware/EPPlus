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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    internal class SumIfCompiler : FunctionCompiler
    {
        public SumIfCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }

        private readonly ExpressionEvaluator _evaluator = new ExpressionEvaluator();

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            if(children.Count() == 3 && children.ElementAt(2).HasChildren)
            {
                var lastExp = children.ElementAt(2).Children.First();
                lastExp.IgnoreCircularReference = true;
                var currentAdr = Context.Scopes.Current.Address;
                var sumRangeAdr = new ExcelAddress(lastExp.ExpressionString);
                var sumRangeWs = string.IsNullOrEmpty(sumRangeAdr.WorkSheetName) ? currentAdr.Worksheet : sumRangeAdr.WorkSheetName;
                if(currentAdr.Worksheet == sumRangeWs && sumRangeAdr.Collide(new ExcelAddress(currentAdr.Address)) != ExcelAddressBase.eAddressCollition.No)
                {
                    var candidateArg = children.ElementAt(1)?.Children.FirstOrDefault()?.Compile().Result;
                    if(children.ElementAt(0).HasChildren)
                    {
                        var functionRowIndex = (currentAdr.FromRow - sumRangeAdr._fromRow);
                        var functionColIndex = (currentAdr.FromCol - sumRangeAdr._fromCol);
                        var firstRangeResult = children.ElementAt(0).Children.First().Compile().Result as IRangeInfo;
                        if(firstRangeResult != null)
                        {
                            var candidateRowIndex = firstRangeResult.Address._fromRow + functionRowIndex;
                            var candidateColIndex = firstRangeResult.Address._fromCol + functionColIndex;
                            var candidateValue = firstRangeResult.GetValue(candidateRowIndex, candidateColIndex);
                            if(_evaluator.Evaluate(candidateArg, candidateValue.ToString()))
                            {
                                if(Context.Configuration.AllowCircularReferences)
                                {
                                    return CompileResult.ZeroDecimal;
                                }
                                throw new CircularReferenceException("Circular reference detected in " + currentAdr.Address);
                            }
                        }
                        
                    }
                }
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
