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
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of cells (of a supplied range), that satisfy a given criteria")]
    internal class CountIf : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return new ArrayBehaviourConfig { ArrayParameterIndexes = new List<int> { 1 } };
        }
        private ExpressionEvaluator _expressionEvaluator;
        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression, false);
            }
            return _expressionEvaluator.Evaluate(obj, expression, false);
        }

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            _expressionEvaluator = new ExpressionEvaluator(context);
            var range = arguments[0];
            var criteria = arguments[1].ValueFirstString;
            double result = 0d;
            if (range.IsExcelRange)
            {
                var rangeInfo = range.ValueAsRangeInfo;
                int fromRow, toRow,fromCol, toCol;
                if(rangeInfo.Address==null)
                {
                    fromRow = fromCol = 0;
                    toRow = rangeInfo.Size.NumberOfRows-1;
                    toCol = rangeInfo.Size.NumberOfCols-1;
                }
                else
                {
                    fromRow = rangeInfo.Address.FromRow;
                    toRow = rangeInfo.Address.ToRow;
                    fromCol = rangeInfo.Address.FromCol;
                    toCol = rangeInfo.Address.ToCol;
                }
                for (int row = fromRow; row <= toRow; row++)
                {
                    for (int col = fromCol; col <= toCol; col++)
                    {
                        if (criteria != null && Evaluate(rangeInfo.GetValue(row, col), criteria))
                        {
                            result++;
                        }
                    }
                }
            }
            else if (range.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var arg in (IEnumerable<FunctionArgument>) range.Value)
                {
                    if(Evaluate(arg.Value, criteria))
                    {
                        result++;
                    }
                }
            }
            else
            {
                if (Evaluate(range.Value, criteria))
                {
                    result++;
                }
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
