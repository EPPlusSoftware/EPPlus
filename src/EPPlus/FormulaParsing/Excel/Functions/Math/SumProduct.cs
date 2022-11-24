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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the sum of the products of corresponding values in two or more supplied arrays")]
    internal class SumProduct : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double result = 0d;
            List<List<double>> results = new List<List<double>>();
            foreach (var arg in arguments)
            {
                results.Add(new List<double>());
                var currentResult = results.Last();
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    foreach (var val in (IEnumerable<FunctionArgument>)arg.Value)
                    {
                        AddValue(val.Value, currentResult);
                    }
                }
                else if (arg.Value is FunctionArgument)
                {
                    AddValue(arg.Value, currentResult);
                }
                else if (arg.IsExcelRange)
                {
                    var r=arg.ValueAsRangeInfo;
                    for (int col = r.Address.FromCol; col <= r.Address.ToCol; col++)
                    {
                        for (int row = r.Address.FromRow; row <= r.Address.ToRow; row++)
                        {
                            AddValue(r.GetValue(row,col), currentResult);
                        }
                    }
                }
                else if(IsNumeric(arg.Value))
                {
                    AddValue(arg.Value, currentResult);
                }
            }
            // Validate that all supplied lists have the same length
            var arrayLength = results.First().Count;
            foreach (var list in results)
            {
                if (list.Count != arrayLength)
                {
                    throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
                    //throw new ExcelFunctionException("All supplied arrays must have the same length", ExcelErrorCodes.Value);
                }
            }
            for (var rowIndex = 0; rowIndex < arrayLength; rowIndex++)
            {
                double rowResult = 1;
                for (var colIndex = 0; colIndex < results.Count; colIndex++)
                {
                    rowResult *= results[colIndex][rowIndex];
                }
                result += rowResult;
            }
            return CreateResult(result, DataType.Decimal);
        }

        private void AddValue(object convertVal, List<double> currentResult)
        {
            if (IsNumeric(convertVal))
            {
                currentResult.Add(Convert.ToDouble(convertVal));
            }
            else if (convertVal is ExcelErrorValue)
            {
                throw (new ExcelErrorValueException((ExcelErrorValue)convertVal));
            }
            else
            {
                currentResult.Add(0d);
            }
        }
    }
}
