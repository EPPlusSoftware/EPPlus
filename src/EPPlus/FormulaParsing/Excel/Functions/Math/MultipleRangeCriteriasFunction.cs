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
    public abstract class MultipleRangeCriteriasFunction : ExcelFunction
    {

        private readonly ExpressionEvaluator _expressionEvaluator;

        protected MultipleRangeCriteriasFunction()
            :this(new ExpressionEvaluator())
        {
            
        }

        protected MultipleRangeCriteriasFunction(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _expressionEvaluator = evaluator;
        }

        protected bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _expressionEvaluator.Evaluate(obj, expression);
        }

        protected List<int> GetMatchIndexes(ExcelDataProvider.IRangeInfo rangeInfo, string searched)
        {
            var result = new List<int>();
            var internalIndex = 0;
            var toRow = rangeInfo.Address._toRow;
            if(rangeInfo.Worksheet.Dimension.End.Row < toRow)
            {
                toRow = rangeInfo.Worksheet.Dimension.End.Row;
            }
            for (var row = rangeInfo.Address._fromRow; row <= toRow; row++)
            {
                if(row > 5)
                {
                    int i = 0;
                }
                for (var col = rangeInfo.Address._fromCol; col <= rangeInfo.Address._toCol; col++)
                {
                    var candidate = rangeInfo.GetValue(row, col);
                    if (searched != null && Evaluate(candidate, searched))
                    {
                        result.Add(internalIndex);
                    }
                    internalIndex++;
                }
            }
            return result;
        }
    }
}
