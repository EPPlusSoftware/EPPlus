/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the probability that values in a range are between two limits")]

    internal class Prob : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count > 4) return CompileResult.GetErrorResult(eErrorType.Value);

            List<double?> xRange;
            if (arguments[0].IsExcelRange)
            {
                xRange = RangeFlattener.FlattenRange(arguments[0].ValueAsRangeInfo, true);
            }
            else
            {
                var a1a = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
                xRange = new List<double?> { a1a };
            }
            List<double?> probRange;
            if (arguments[1].IsExcelRange)
            {
                probRange = RangeFlattener.FlattenRange(arguments[1].ValueAsRangeInfo, true);
            }
            else
            {
                var a1b = ArgToDecimal(arguments, 1, out ExcelErrorValue e3);
                if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
                probRange = new List<double?> { a1b };
            }

            if (xRange.Count != probRange.Count)
            {
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            if (probRange.SumKahan() != 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var lowerLimit = ArgToDecimal(arguments, 2, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            double upperLimit;
            if (arguments.Count < 4)
            {
                upperLimit = lowerLimit;
            }
            else
            {
                upperLimit = ArgToDecimal(arguments, 3, out ExcelErrorValue e5);
                if(e5 != null) return CompileResult.GetErrorResult(e5.Type);
            }

            var result = 0d;
            for (var i = 0; i < xRange.Count; i++)
            {
                var x = xRange[i];
                var prob = probRange[i];
                if (x >= lowerLimit && x <= upperLimit)
                {
                    result += prob ?? 0;
                }
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
