/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "5.5",
            IntroducedInExcelVersion = "2010",
            Description = "Returns the specified quartile of a set of supplied numbers, based on percentile value 0 - 1 (exclusive) ")]
    internal class QuartileExc : PercentileExc
    {
        public override string NamespacePrefix => "_xlfn.";
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arrArg = arguments.Take(1);
            var arr = ArgsToDoubleEnumerable(arrArg, context, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            if (!arr.Any()) return CreateResult(eErrorType.Value);
            var quart = ArgToInt(arguments, 1);
            switch (quart)
            {
                case 1:
                    return base.Execute(BuildArgs(arrArg, 0.25d), context);
                case 2:
                    return base.Execute(BuildArgs(arrArg, 0.5d), context);
                case 3:
                    return base.Execute(BuildArgs(arrArg, 0.75d), context);
                default:
                    return CreateResult(eErrorType.Num);
            }
        }

        private IList<FunctionArgument> BuildArgs(IEnumerable<FunctionArgument> arrArg, double quart)
        {
            var argList = new List<FunctionArgument>();
            argList.AddRange(arrArg);
            argList.Add(new FunctionArgument(quart, DataType.Decimal));
            return argList;
        }
    }
}
