using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7.0",
        Description = "Returns test for independence with help of chi-square statistic and appropriate degrees of freedom.")]
    internal class ChisqTest : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // both arguments should be ranges (with more than 1 cell).
            if (!arguments[0].IsExcelRange || !arguments[1].IsExcelRange)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }

            var observed = ArgToRangeInfo(arguments, 0);
            var expected = ArgToRangeInfo(arguments, 1);
            // Chi-squared function, two ranges as argument and returns the independence level for the two ranges.

            double chisq = 0d;
            double df = 0d;

            if (observed.Size.NumberOfRows == 1 && observed.Size.NumberOfCols == 1)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            else if (observed.Size.NumberOfRows != expected.Size.NumberOfRows || observed.Size.NumberOfCols != expected.Size.NumberOfCols)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            else
            {
                if (observed.Size.NumberOfRows > 1 && observed.Size.NumberOfCols > 1)
                {
                    df = (observed.Size.NumberOfRows - 1) / (observed.Size.NumberOfCols - 1);
                }
                else if (observed.Size.NumberOfRows == 1 && observed.Size.NumberOfCols > 1)
                {
                    df = observed.Size.NumberOfCols - 1;
                }
                else if (observed.Size.NumberOfRows > 1 && observed.Size.NumberOfCols == 1)
                {
                    df = observed.Size.NumberOfRows - 1;
                }
            }
            for (var i = 0; i < observed.Size.NumberOfRows; i++) 
            {
                for (var j = 0; j < observed.Size.NumberOfCols; j++)
                {
                    var v1 = observed.GetOffset(i, j);
                    var doubleV1 = ConvertUtil.GetValueDouble(v1);
                    var v2 = expected.GetOffset(i, j);
                    var doubleV2 = ConvertUtil.GetValueDouble(v2);
                    chisq += System.Math.Pow(doubleV1 - doubleV2, 2) / (doubleV2);
                }
            }

            var result = 1d - ChiSquareHelper.CumulativeDistribution(chisq, df);

            return CreateResult(result, DataType.Decimal);
        }
    }
}
