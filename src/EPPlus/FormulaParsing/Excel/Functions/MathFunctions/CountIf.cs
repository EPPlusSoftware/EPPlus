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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(1);
        }
        private ExpressionEvaluator _expressionEvaluator;
        private bool Evaluate(object obj, string expression)
        {
            if(expression==null)
            {
                expression = "0";
            }
            if (IsNumeric(obj))
            {                
                return _expressionEvaluator.Evaluate(ConvertUtil.GetValueDouble(obj), expression, true);
            }
            return _expressionEvaluator.Evaluate(obj, expression, false);
        }
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            _expressionEvaluator = new ExpressionEvaluator(context);
            var range = arguments[0];
            var arg1 = arguments[1].ValueFirst??0;
            string criteria = null;
            bool isString = false;
            bool isEmptyCriteria = false;
            if (arg1 is string s)
            {
                criteria = s;
                isString = true;
                isEmptyCriteria = criteria == string.Empty || criteria.Trim() == "=";
            }
            double result = 0d;
            if (range.IsExcelRange)
            {
                var rangeInfo = range.ValueAsRangeInfo;
                if (rangeInfo.Address.FromRow <= 0)
                {
                    var toRow = rangeInfo.Size.NumberOfRows;
                    var toCol = rangeInfo.Size.NumberOfCols;
                    for (int r = 0; r < toRow; r++)
                    {
                        for (int c = 0; c < toCol; c++)
                        {
                            var v = rangeInfo.GetValue(r, c);
                            if (isString)
                            {
                                if (Evaluate(v, criteria))
                                {
                                    result++;
                                }
                            }
                            else
                            {
                                if (ConvertUtil.GetValueDouble(v) == ConvertUtil.GetValueDouble(arg1))
                                {
                                    result++;
                                }
                            }
                        }
                    }
                }
                else
                {
                    var emptyCells = 0;
                    var cse = new CellStoreEnumerator<ExcelValue>(rangeInfo.Worksheet._values, rangeInfo.Address.ToExcelAddressBase());
                    int row = range.Address.FromRow;
                    int col = range.Address.FromCol;
                    int add = 1;
                    foreach (var c in cse)
                    {
                        if (isEmptyCriteria)
                        {
                            emptyCells += CalculateEmptyCells(row, col, cse.Row, cse.Column, cse, add);
                        }

                        row = cse.Row;
                        col = cse.Column;
                        if (isString)
                        {
                            if (Evaluate(cse.Value._value, criteria))
                            {
                                result++;
                            }
                        }
                        else
                        {
                            var v = cse.Value._value;
                            if (string.IsNullOrEmpty(v?.ToString())==false && ConvertUtil.GetValueDouble(v) == ConvertUtil.GetValueDouble(arg1))
                            {
                                result++;
                            }
                        }
                        add = 0;
                    }

                    if (isEmptyCriteria) //Check for null values
                    {
                        result += emptyCells + CalculateEmptyCells(row, col, cse._endRow, cse._endCol, cse, 1);
                    }
                }
            }
            else if (range.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var arg in (IEnumerable<FunctionArgument>)range.Value)
                {
                    if (isString)
                    {
                        if (Evaluate(arg.Value, criteria))
                        {
                            result++;
                        }
                    }
                    else
                    {
                        if (ConvertUtil.GetValueDouble(arg.Value) == ConvertUtil.GetValueDouble(arg1))
                        {
                            result++;
                        }
                    }
                }
            }
            else
            {
                if(isString)
                { 
                    if (Evaluate(range.Value, criteria))
                    {
                        result++;
                    }
                }
                else
                {
                    if (ConvertUtil.GetValueDouble(range.Value) == ConvertUtil.GetValueDouble(arg1))
                    {
                        result++;
                    }
                }
            }
            return CreateResult(result, DataType.Integer);
        }

        private int CalculateEmptyCells(int row, int col, int nextRow, int nextCol, CellStoreEnumerator<ExcelValue> cse, int add)
        {
            if(row == nextRow)
            {
                if(col < cse._endCol && col < nextCol+1)
                {
                    return nextCol - col;
                }
                return 0;
            }
            else
            {
                var rows = nextRow - row-1;
                var cells = cse._endCol - col + add;               //Add a cell for the first item in a range. 
                cells += nextCol - cse._startCol;
                cells += rows * (cse._endCol - cse._startCol + 1);
                return cells;
            }
        }

        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 1)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            return FunctionParameterInformation.Normal;
        }));
    }
}
