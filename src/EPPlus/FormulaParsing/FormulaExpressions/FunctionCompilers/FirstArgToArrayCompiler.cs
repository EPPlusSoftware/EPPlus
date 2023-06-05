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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class FirstArgToArrayCompiler : FunctionCompiler
    {
        internal FirstArgToArrayCompiler(ExcelFunction function, ParsingContext context)
            : this(function, context, false)
        {

        }

        internal FirstArgToArrayCompiler(ExcelFunction function, ParsingContext context, bool handleErrors)
            : base(function, context)
        {
            _handleErrors = handleErrors;
        }

        private readonly bool _handleErrors;

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            if (!children.Any()) return new CompileResult(eErrorType.Value);
            var firstChild = children.First();
            var compileResult = firstChild.Compile();
            if(compileResult.DataType == DataType.ExcelRange)
            {
                var remainingChildren = children.Skip(1).ToList();
                var range = compileResult.Result as IRangeInfo;
                if(range.Size.NumberOfCols > 1 || range.Size.NumberOfRows > 1)
                {
                    var rangeDef = new RangeDefinition(range.Size.NumberOfRows, range.Size.NumberOfCols);
                    var inMemoryRange = new InMemoryRange(rangeDef);
                    var errorCompileResult = default(CompileResult);
                    for(var row = 0; row < rangeDef.NumberOfRows; row++)
                    {
                        errorCompileResult = default(CompileResult);
                        for (var col = 0; col < rangeDef.NumberOfCols; col++)
                        {
                            errorCompileResult = default(CompileResult);
                            var argAsCompileResult = CompileResultFactory.Create(range.GetOffset(row, col));
                            var arg = new FunctionArgument(argAsCompileResult.ResultValue, argAsCompileResult.DataType);
                            var argList = new List<FunctionArgument> { arg };
                            argList.AddRange(remainingChildren.Select(x =>
                            {
                                if(_handleErrors)
                                {
                                    try
                                    {
                                        var cr = x.Compile();
                                        return new FunctionArgument(cr.ResultValue, cr.DataType);
                                    }
                                    catch (ExcelErrorValueException efe)
                                    {
                                        errorCompileResult = ((ErrorHandlingFunction)Function).HandleError(efe.ErrorValue.ToString());
                                        return null;
                                    }
                                    catch// (Exception e)
                                    {
                                        errorCompileResult = ((ErrorHandlingFunction)Function).HandleError(ExcelErrorValue.Values.Value);
                                        return null;
                                    }
                                }
                                else
                                {
                                    var cr = x.Compile();
                                    return new FunctionArgument(cr.Result, cr.DataType);
                                }
                            }));
                            if(errorCompileResult != null)
                            {
                                inMemoryRange.SetValue(row, col, errorCompileResult.Result);
                            }
                            else
                            {
                                var result = Function.ExecuteInternal(argList, Context);
                                inMemoryRange.SetValue(row, col, result.Result);
                            }
                        }
                    }
                    return new CompileResult(inMemoryRange, DataType.ExcelRange);
                }
                else
                {
                    var defaultCompiler = new DefaultCompiler(Function, Context);
                    return defaultCompiler.Compile(children);
                }
                
            }
            else
            {
                var defaultCompiler = new DefaultCompiler(Function, Context);
                return defaultCompiler.Compile(children);
            }
        }
    }
}
