using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "5.5",
            Description = "Returns the K'th percentile of values in a supplied range, where K is in the range 0 - 1 (inclusive)")]
    internal class QuartileInc : PercentileInc
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arrArg = arguments.Take(1);
            var arr = ArgsToDoubleEnumerable(arrArg, context).Select(x => (double)x).ToList();
            if (!arr.Any()) return CreateResult(eErrorType.Value);
            var quart = ArgToInt(arguments, 1);
            switch (quart)
            {
                case 0:
                    return CreateResult(arr.Min(), DataType.Decimal);
                case 1:
                    return base.Execute(BuildArgs(arrArg, 0.25d), context);
                case 2:
                    return base.Execute(BuildArgs(arrArg, 0.5d), context);
                case 3:
                    return base.Execute(BuildArgs(arrArg, 0.75d), context);
                case 4:
                    return CreateResult(arr.Max(), DataType.Decimal);
                default:
                    return CreateResult(eErrorType.Num);
            }
        }

        private IEnumerable<FunctionArgument> BuildArgs(IEnumerable<FunctionArgument> arrArg, double quart)
        {
            var argList = new List<FunctionArgument>();
            argList.AddRange(arrArg);
            argList.Add(new FunctionArgument(quart, DataType.Decimal));
            return argList;
        }
    }
}
