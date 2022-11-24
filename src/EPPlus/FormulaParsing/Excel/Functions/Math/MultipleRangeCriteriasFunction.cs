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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal abstract class MultipleRangeCriteriasFunction : ExcelFunction
    {
        protected bool Evaluate(object obj, string expression, ParsingContext ctx, bool convertNumericString = true)
        {
            var expressionEvaluator = new ExpressionEvaluator(ctx);
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return expressionEvaluator.Evaluate(candidate.Value, expression, convertNumericString);
            }
            return expressionEvaluator.Evaluate(obj, expression, convertNumericString);
        }

        protected List<int> GetMatchIndexes(RangeOrValue rangeOrValue, string searched, ParsingContext ctx, bool convertNumericString = true)
        {
            var expressionEvaluator = new ExpressionEvaluator(ctx);
            var result = new List<int>();
            var internalIndex = 0;
            if (rangeOrValue.Range != null)
            {
                var rangeInfo = rangeOrValue.Range;
                var toRow = rangeInfo.Address.ToRow;
                if (rangeInfo.Worksheet.Dimension.End.Row < toRow)
                {
                    toRow = rangeInfo.Worksheet.Dimension.End.Row;
                }
                for (var row = rangeInfo.Address.FromRow; row <= toRow; row++)
                {
                    for (var col = rangeInfo.Address.FromCol; col <= rangeInfo.Address.ToCol; col++)
                    {
                        var candidate = rangeInfo.GetValue(row, col);
                        if (searched != null && Evaluate(candidate, searched, ctx, convertNumericString))
                        {
                            result.Add(internalIndex);
                        }
                        internalIndex++;
                    }
                }
            }
            else if(Evaluate(rangeOrValue.Value, searched, ctx, convertNumericString))
            {
                result.Add(internalIndex);
            }
            return result;
        }
    }
}
