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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils.TypeConversion;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Adds the cells in a supplied range, that satisfy a given criteria")]
    internal class SumIf : RangeCriteriaFunction
    {
        private ExpressionEvaluator _evaluator;
        public override int ArgumentMinLength => 2;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(1);
        }
        public override void GetNewParameterAddress(IList<CompileResult> args, int index, ref Queue<FormulaRangeAddress> addresses)
        {
            if(index == 2)
            {
                if (args[0].Result is IRangeInfo criteriaRange && args[2].Result is IRangeInfo valueRange)
                {
                    var rv = new RangeOrValue { Range = criteriaRange };                    
                    var  mi=GetMatchIndexes(rv, args[1].Result, null);
                    addresses = EnqueueMatchingAddresses(valueRange, mi, ref addresses);
                }
            }
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            _evaluator = new ExpressionEvaluator(context);
            var argRange = ArgToRangeInfo(arguments, 0);

            // Criteria can either be a string or an array of strings
            var criteria = GetCriteria(arguments.ElementAt(1));
            KahanSum retVal = 0d;
            if (argRange == null)
            {
                var val = arguments[0].Value;
                if (_evaluator.Evaluate(val, criteria))
                {
                    if (arguments.Count > 2)
                    {
                        var sumRange = ArgToRangeInfo(arguments, 2);
                        retVal = sumRange.First().ValueDouble;
                    }
                    else
                    {
                        retVal = ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            else if (arguments.Count > 2)
            {
                var sumRange = ArgToRangeInfo(arguments, 2);
                retVal = CalculateWithSumRange(argRange, criteria, sumRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(argRange, criteria, context);
            }
            return CreateResult(retVal.Get(), DataType.Decimal);
        }

        internal static IEnumerable<string> GetCriteria(FunctionArgument criteriaArg)
        {
            var criteria = new List<string>();

            if (criteriaArg.IsExcelRange)
            {
                foreach (var cell in criteriaArg.ValueAsRangeInfo)
                {
                    if (cell.Value != null)
                    {
                        criteria.Add(cell.Value.ToString());
                    }
                }
            }
            else
            {
                criteria.Add(criteriaArg.ValueFirst != null ? criteriaArg.ValueFirst.ToString() : null);
            }
            return criteria;
        }

        private double CalculateWithSumRange(IRangeInfo range, IEnumerable<string> criteria, IRangeInfo sumRange, ParsingContext context)
        {
            KahanSum retVal = 0d;
            foreach (var cell in range)
            {
                if (_evaluator.Evaluate(cell.Value, criteria))
                {
                    var rowOffset = cell.Row - range.Address.FromRow;
                    var columnOffset = cell.Column - range.Address.FromCol;
                    if (sumRange.Address.FromRow + rowOffset <= sumRange.Address.ToRow &&
                       sumRange.Address.FromCol + columnOffset <= sumRange.Address.ToCol)
                    {
                        var val = sumRange.GetOffset(rowOffset, columnOffset);
                        if (val is ExcelErrorValue)
                        {
                            ThrowExcelErrorValueException((ExcelErrorValue)val);
                        }                        
                        retVal += ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            return retVal.Get();
        }

        //Dictionary<IRangeInfo, IEnumerable<string>> rangeExpressionsCache;
        //Dictionary<IRangeInfo, Dictionary<string, double>> rangeValues;

        private double CalculateSingleRange(IRangeInfo range, IEnumerable<string> expressions, ParsingContext context)
        {
            KahanSum retVal = 0d;
            var rangeValues = context.rangeValues;
            if (rangeValues == null)
            {
                context.rangeValues = new();
                rangeValues = context.rangeValues;
            }

            if (rangeValues.TryGetValue(range.Address.WorksheetAddress, out Dictionary<string, double> valueDict) == false)
            {
                valueDict = new Dictionary<string, double>();
                foreach (var candidate in range)
                {
                    if (IsNumeric(candidate.Value))
                    {
                        if (candidate.IsExcelError)
                        {
                            ThrowExcelErrorValueException((ExcelErrorValue)candidate.Value);
                        }
                        valueDict.Add(candidate.Address, candidate.ValueDouble);

                        if (_evaluator.Evaluate(candidate.Value, expressions))
                        {
                            retVal += candidate.ValueDouble;
                        }
                    }
                }
                context.rangeValues.Add(range.Address.WorksheetAddress, valueDict);
            }
            else
            {
                foreach (var key in valueDict.Keys)
                {
                    if (_evaluator.Evaluate(valueDict[key], expressions))
                    {
                        retVal += valueDict[key];
                    }
                }
            }
            return retVal.Get();
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 1)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            else if(argumentIndex == 2)
            {
                return FunctionParameterInformation.AdjustParameterAddress;
            }
            return FunctionParameterInformation.Normal;
        }));
        public override bool IsVolatile => true;
    }
}
