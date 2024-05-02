/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.MathAndTrig,
    EPPlusVersion = "7.2",
    Description = "Get the inverse of Matrix",
    SupportsArrays = true)]
    internal class MUnit : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].IsExcelRange && arguments[0].Value is IRangeInfo range)
            {
                var returnRange = new InMemoryRange(range.Size);
                for (int r = 0; r < range.Size.NumberOfRows; r++)
                {
                    for (int c = 0; c < range.Size.NumberOfCols; c++)
                    {
                        var v = range.GetOffset(r, c);
                        if (ConvertUtil.IsExcelNumeric(v) == false || DoubleArgParser.Parse(v, out ExcelErrorValue _) <= 0d)
                        {
                            returnRange.SetValue(r, c, ErrorValues.ValueError);
                        }
                        else
                        {
                            returnRange.SetValue(r, c, 1d);
                        }
                    }
                }
                return CreateDynamicArrayResult(returnRange, DataType.ExcelRange);
            }
            if (ConvertUtil.IsExcelNumeric(arguments[0].ValueFirst) == false )
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            var size = ArgToInt(arguments, 0, RoundingMethod.Convert);
            if (size <= 0)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            double[][] m = MatrixHelper.GetIdentityMatrix(size);
            var returnRange2 = new InMemoryRange(size, (short)size);
            for(int i = 0; i< m.Length;i++)
            {
                for(int j=0; j< m[i].Length; j++)
                {
                    returnRange2.SetValue(i, j, m[i][j]);
                }
            }
            return CreateDynamicArrayResult(returnRange2, DataType.ExcelRange);
        }
    }
}
