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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class CustomArrayBehaviourCompiler : FunctionCompiler
    {
        internal CustomArrayBehaviourCompiler(ExcelFunction function, ParsingContext context)
            : this(function, context, false)
        {

        }

        internal CustomArrayBehaviourCompiler(ExcelFunction function, ParsingContext context, bool handleErrors)
            : base(function)
        {
            _handleErrors = handleErrors;
        }

        private readonly bool _handleErrors;

        public override CompileResult Compile(IEnumerable<CompileResult> children, ParsingContext context)
        {
            var args = new List<FunctionArgument>();

            if (!children.Any()) return new CompileResult(eErrorType.Value);

            var rangeArgs = new Dictionary<int, IRangeInfo>();
            var otherArgs = new Dictionary<int, CompileResult>();
            for (var ix = 0; ix < children.Count(); ix++)
            {
                var cr = children.ElementAt(ix);
                if (cr.DataType == DataType.ExcelRange && Function.ArrayBehaviourConfig.CanBeArrayArg(ix) && (cr.Result is IRangeInfo || cr.Result is FormulaRangeAddress))
                {
                    var range = cr.Result as IRangeInfo;
                    if (range == null && cr.Result is FormulaRangeAddress fra)
                    {
                        range = new RangeInfo(fra);
                    }
                    rangeArgs[ix] = range;

                    //OLD CODE left below in case otherArgs becomes a problem later. Range args when NCells == 1 is
                    //neccesary in case of in memory range/dynamic array formulas with a single cell.

                    //if (range.GetNCells() > 1)
                    //{
                    //    rangeArgs[ix] = range;
                    //}
                    //else
                    //{
                    //    otherArgs[ix] = CompileResultFactory.Create(range.GetOffset(0, 0));
                    //}
                }
                else
                {
                    if (cr.DataType == DataType.ExcelRange)
                    {
                        cr = CompileResultFactory.Create(cr.Result);
                    }
                    otherArgs[ix] = cr;
                }
            }

            if (rangeArgs.Count == 0)
            {
                var defaultCompiler = new DefaultCompiler(Function);
                return defaultCompiler.Compile(children, context);
            }

            short maxWidth = 0;
            var maxPhysicalHeight = 0;
            var maxLogicalHeight = 0;
            foreach (var rangeArg in rangeArgs.Values)
            {
                if (rangeArg.Size.NumberOfCols > maxWidth)
                {
                    maxWidth = rangeArg.Size.NumberOfCols;
                }

                int physical = RangeHelper.GetPhysicalRows(rangeArg);
                if (physical > maxPhysicalHeight)
                {
                    maxPhysicalHeight = physical;
                }
                if (rangeArg.Size.NumberOfRows > maxLogicalHeight)
                {
                    maxLogicalHeight = rangeArg.Size.NumberOfRows;
                }
            }

            InMemoryRange resultRange;
            if (maxPhysicalHeight < maxLogicalHeight)
            {
                var physicalDef = new RangeDefinition(maxPhysicalHeight, maxWidth);
                if (rangeArgs.Count == 1)
                {
                    resultRange = new InMemoryRange(rangeArgs.First().Value.Address, physicalDef, maxLogicalHeight);
                }
                else
                {
                    resultRange = new InMemoryRange(physicalDef, maxLogicalHeight);
                }
            }
            else
            {
                var rangeDef = new RangeDefinition(maxLogicalHeight, maxWidth);
                if (rangeArgs.Count == 1)
                {
                    resultRange = new InMemoryRange(rangeArgs.First().Value.Address, rangeDef);
                }
                else
                {
                    resultRange = new InMemoryRange(rangeDef);
                }
            }

            var nArgs = children.Count();
            for (var row = 0; row < resultRange.PhysicalRows; row++)
            {
                for (var col = 0; col < resultRange.Size.NumberOfCols; col++)
                {
                    bool isError = false;
                    var argList = new List<FunctionArgument>();
                    for (var argIx = 0; argIx < nArgs; argIx++)
                    {
                        if (rangeArgs.ContainsKey(argIx))
                        {
                            var range = rangeArgs[argIx];
                            var r = row;
                            var c = col;
                            if (range.Size.NumberOfCols == 1 && range.Size.NumberOfRows == resultRange.Size.NumberOfRows)
                            {
                                c = 0;
                            }
                            if (range.Size.NumberOfRows == 1 && range.Size.NumberOfCols == resultRange.Size.NumberOfCols)
                            {
                                r = 0;
                            }
                            if (col < range.Size.NumberOfCols && row < range.Size.NumberOfRows || (c != col) || (r != row))
                            {
                                var argAsCompileResult = CompileResultFactory.Create(range.GetOffset(r, c));
                                argList.Add(new FunctionArgument(argAsCompileResult));
                            }
                            else
                            {
                                if (Function.IsErrorHandlingFunction)
                                {
                                    argList.Add(new FunctionArgument(ErrorValues.NAError));
                                }
                                else
                                {
                                    resultRange.SetValue(row, col, ErrorValues.NAError);
                                    isError = true;
                                    continue;
                                }
                            }

                        }
                        else
                        {
                            var arg = otherArgs[argIx];
                            if (arg.DataType == DataType.LambdaCalculation)
                            {
                                var calculator = arg.ResultValue as LambdaCalculator;
                                calculator.BeginCalculation();
                                calculator.ResetVariables();
                            }
                            argList.Add(new FunctionArgument(arg));
                        }
                    }
                    if (!isError)
                    {
                        var result = Function.ExecuteInternal(argList, context);
                        resultRange.SetValue(row, col, result.Result);
                    }
                }
            }
            if (resultRange.Address?.FromRow > 0 && Function.ReturnsReference)
            {
                return new AddressCompileResult(resultRange, DataType.ExcelRange, resultRange.Address);
            }
            else
            {
                return new CompileResult(resultRange, DataType.ExcelRange);
            }
        }
    }
}