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
            if (criteria.Any(x=>x==""))
            {
                var adrDA = range.GetAddressDimensionAdjusted(0).Address;
                for(int r=adrDA.FromRow; r <= adrDA.ToRow; r++)
                {
                    for(int c= adrDA.FromCol; c <= adrDA.ToCol; c++)
                    {
                        CalculateCell(range, criteria, sumRange, r, c, range.GetValue(r,c), ref retVal);
                    }
                }
            }
            else
            {
                foreach (var cell in range)
                {
                    CalculateCell(range, criteria, sumRange, cell.Row, cell.Column, cell.Value, ref retVal);
                }
            }
            return retVal.Get();
        }

        private void CalculateCell(IRangeInfo range, IEnumerable<string> criteria, IRangeInfo sumRange, int row, int col, object value, ref KahanSum retVal)
        {
            if (_evaluator.Evaluate(value, criteria))
            {
                var rowOffset = row - range.Address.FromRow;
                var columnOffset = col - range.Address.FromCol;
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

        private double CalculateSingleRange(IRangeInfo range, IEnumerable<string> expressions, ParsingContext context)
        {
            KahanSum retVal = 0d;
            foreach (var candidate in range)
            {
                if (IsNumeric(candidate.Value) && _evaluator.Evaluate(candidate.Value, expressions))
                {
                    if (candidate.IsExcelError)
                    {
                        ThrowExcelErrorValueException((ExcelErrorValue)candidate.Value);
                    }
                    retVal += candidate.ValueDouble;
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
