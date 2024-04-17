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
using System.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "7.2",
        Description = "Multiply to matrixes",
        SupportsArrays = true)]
    internal class MInverse : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = ArgToRangeInfo(arguments, 0);
            var r = range.Address.ToRow - range.Address.FromRow + 1;
            var c = range.Address.ToCol - range.Address.FromCol + 1;
            if (r != c)
            {
                return CreateResult(CompileResult.GetErrorResult(eErrorType.Value), DataType.ExcelRange);
            }
            double[][] m = new double[r][];
            for (int i = 0; i < r; i++)
            {
                m[i] = new double[c];
            }
            var x = range.Address.FromCol;
            var y = range.Address.FromRow;
            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
                {
                    bool e1 = double.TryParse(range.GetValue(y, x).ToString(), out double t);
                    if( !e1 )
                    {
                        return CreateResult(CompileResult.GetErrorResult(eErrorType.Value), DataType.ExcelRange);
                    }
                    m[i][j] = t;
                    x++;
                }
                x = 1;
                y++;
            }
            //if determinant != 0
            var returnRange = CompileResult.GetErrorResult(eErrorType.Value);
            //calc inverse
            //populate returnRange from double[][]
            return CreateResult(returnRange, DataType.ExcelRange);
        }
    }
}
