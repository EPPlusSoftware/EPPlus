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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "7.2",
        Description = "Multiply to matrixes",
        SupportsArrays = true)]
    internal class MMult : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var m1 = ArgToRangeInfo(arguments, 0);
            var m2 = ArgToRangeInfo(arguments, 1);
            var r1 = m1.Address.ToRow - m1.Address.FromRow;
            var c1 = m1.Address.ToCol - m1.Address.FromCol;
            var r2 = m2.Address.ToRow - m2.Address.FromRow;
            var c2 = m2.Address.ToCol - m2.Address.FromCol;
            if(c1 != r2)
            {
                return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
            }
            double temp = 0;
            double[,] result = new double[r1+1, c2+1];
            var returnRange = new InMemoryRange(r1 + 1, (short)(c2 + 1));
            int x1 = m1.Address.FromRow, y1 = m1.Address.FromCol;
            int x2 = m2.Address.FromRow, y2 = m2.Address.FromCol;

            for (int i = 0; i <= r1; i++)
            {
                for (int j = 0; j <= c2; j++)
                {
                    temp = 0;
                    for (int k = 0; k <= c1; k++)
                    {
                        var e1 = m1.GetValue(x1, y1);
                        var e2 = m2.GetValue(x2, y2);

                        if( !IsNumeric(e1) || !IsNumeric(e2))
                        {
                            return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                        }
                        var t1 = ConvertUtil.GetValueDouble(e1);
                        var t2 = ConvertUtil.GetValueDouble(e2);
                        temp += t1 * t2;
                        y1++;
                        x2++;
                    }
                    returnRange.SetValue(i, j, temp);
                    y1 = m1.Address.FromCol;
                    x2 = m2.Address.FromRow;
                    y2++;
                }
                x1++;
                y2 = m2.Address.FromCol;
            }
            return CreateResult(returnRange, DataType.ExcelRange);
        }
    }
}