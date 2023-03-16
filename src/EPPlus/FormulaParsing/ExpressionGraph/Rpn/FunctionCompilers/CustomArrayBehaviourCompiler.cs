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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    internal class CustomArrayBehaviourCompiler : RpnFunctionCompiler
    {
        internal CustomArrayBehaviourCompiler(ExcelFunction function, ParsingContext context)
            : this(function, context, false)
        {

        }

        internal CustomArrayBehaviourCompiler(ExcelFunction function, ParsingContext context, bool handleErrors)
            : base(function, context)
        {
            _handleErrors = handleErrors;
        }

        private readonly bool _handleErrors;

        public override CompileResult Compile(IEnumerable<RpnExpression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var arrayConfig = Function.GetArrayBehaviourConfig();
            if (arrayConfig == null) throw new InvalidOperationException("If a function is configured to use custom array behaviour it must return a configuration via the GetArrayBehaviour method (overload from ExcelFunction).");
            
            if (!children.Any()) return new CompileResult(eErrorType.Value);

            var rangeArgs = new Dictionary<int, IRangeInfo>();
            var otherArgs = new Dictionary<int, CompileResult>();
            for(var ix = 0; ix < children.Count(); ix++)
            {
                var child = children.ElementAt(ix);
                var cr = child.Compile();
                if(cr.DataType == DataType.ExcelRange && arrayConfig.ArrayParameterIndexes.Contains(ix))
                {
                    var range = cr.Result as IRangeInfo;
                    if(range.IsMulti)
                    {
                        rangeArgs[ix] = range;
                    }
                    else
                    {
                        otherArgs[ix] = CompileResultFactory.Create(range.GetOffset(0, 0));
                    }
                }
                else
                {
                    otherArgs[ix] = cr;
                }
            }

            if(rangeArgs.Count == 0)
            {
                var defaultCompiler = new RpnDefaultCompiler(Function, Context);
                return defaultCompiler.Compile(children);
            }

            short maxWidth = 0;
            var maxHeight = 0;
            foreach(var rangeArg in rangeArgs.Values)
            {
                if(rangeArg.Size.NumberOfCols > maxWidth)
                {
                    maxWidth = rangeArg.Size.NumberOfCols;
                }
                if(rangeArg.Size.NumberOfRows > maxHeight) 
                {
                    maxHeight= rangeArg.Size.NumberOfRows;
                }
            }

            var resultRangeDef = new RangeDefinition(maxHeight, maxWidth);
            var resultRange = new InMemoryRange(resultRangeDef);
            var nArgs = children.Count();
            for(var row = 0; row < resultRange.Size.NumberOfRows; row++)
            {
                for(var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    var argList = new List<FunctionArgument>();
                    for(var argIx = 0; argIx < nArgs; argIx++)
                    {
                        if(rangeArgs.ContainsKey(argIx))
                        {
                            var range = rangeArgs[argIx];
                            if(col < range.Size.NumberOfCols && row < range.Size.NumberOfRows)
                            {
                                var argAsCompileResult = CompileResultFactory.Create(range.GetOffset(row, col));
                                argList.Add(new FunctionArgument(argAsCompileResult.ResultValue, argAsCompileResult.DataType));
                            }
                            else
                            {
                                resultRange.SetValue(row, col, ErrorValues.NAError);
                                continue;
                            }
                            
                        }
                        else
                        {
                            var arg = otherArgs[argIx];
                            argList.Add(new FunctionArgument(arg.ResultValue, arg.DataType));
                        }
                    }
                    var result = Function.Execute(argList, Context);
                    resultRange.SetValue(row, col, result.Result);
                }
            }
            return new CompileResult(resultRange, DataType.ExcelRange);
        }

        #region Starting code
        private CompileResult StartingCode(IEnumerable<RpnExpression> children)
        {
            var firstChild = children.First();
            var compileResult = firstChild.Compile();
            if (compileResult.DataType == DataType.ExcelRange)
            {
                var remainingChildren = children.Skip(1).ToList();
                var range = compileResult.Result as IRangeInfo;
                if (range.IsMulti)
                {
                    var rangeDef = new RangeDefinition(range.Size.NumberOfRows, range.Size.NumberOfCols);
                    var inMemoryRange = new InMemoryRange(rangeDef);
                    var errorCompileResult = default(CompileResult);
                    for (var row = 0; row < rangeDef.NumberOfRows; row++)
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
                                if (_handleErrors)
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
                                    return new FunctionArgument(cr.ResultValue, cr.DataType);
                                }
                            }));
                            if (errorCompileResult != null)
                            {
                                inMemoryRange.SetValue(row, col, errorCompileResult.Result);
                            }
                            else
                            {
                                var result = Function.Execute(argList, Context);
                                inMemoryRange.SetValue(row, col, result.Result);
                            }
                        }
                    }
                    return new CompileResult(inMemoryRange, DataType.ExcelRange);
                }
                else
                {
                    var defaultCompiler = new RpnDefaultCompiler(Function, Context);
                    return defaultCompiler.Compile(children);
                }

            }
            else
            {
                var defaultCompiler = new RpnDefaultCompiler(Function, Context);
                return defaultCompiler.Compile(children);
            }
        }
        #endregion
    }
}
