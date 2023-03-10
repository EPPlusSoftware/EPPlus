/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/10/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class SingleArgToArrayCompiler : RpnFunctionCompiler
    {
        internal SingleArgToArrayCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {

        }

        public override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            if (!children.Any()) return new CompileResult(eErrorType.Value);
            var firstChild = children.First();
            var compileResult = firstChild.Compile();
            if(compileResult.DataType == DataType.ExcelRange)
            {
                var range = compileResult.Result as IRangeInfo;
                if(range.IsMulti)
                {
                    var rangeDef = new RangeDefinition(range.Size.NumberOfRows, range.Size.NumberOfCols);
                    var inMemoryRange = new InMemoryRange(rangeDef);
                    for(var row = 0; row < rangeDef.NumberOfRows; row++)
                    {
                        for(var col = 0; col < rangeDef.NumberOfCols; col++)
                        {
                            var argAsCompileResult = CompileResultFactory.Create(range.GetOffset(row, col));
                            var arg = new FunctionArgument(argAsCompileResult.ResultValue, argAsCompileResult.DataType);
                            var argList = new List<FunctionArgument> { arg };
                            var result = Function.Execute(argList, Context);
                            inMemoryRange.SetValue(row, col, result.Result);
                        }
                    }
                    return new CompileResult(inMemoryRange, DataType.ExcelRange);
                }
                else
                {
                    var argAsCompileResult = CompileResultFactory.Create(range.GetValue(0, 0));
                    var arg = new FunctionArgument(argAsCompileResult.ResultValue, argAsCompileResult.DataType);
                    var argList = new List<FunctionArgument> { arg };
                    return Function.Execute(argList, Context);
                }
                
            }
            else
            {
                var arg = new FunctionArgument(compileResult.Result, compileResult.DataType);
                var argList = new List<FunctionArgument> { arg };
                return Function.Execute(argList, Context);
            }
        }
    }
}
