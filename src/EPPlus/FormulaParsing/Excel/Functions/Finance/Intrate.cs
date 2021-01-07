using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.5",
        Description = "Calculates the interest rate for a fully invested security")]
    internal class Intrate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var investment = ArgToDecimal(arguments, 2);
            var redemption = ArgToDecimal(arguments, 3);
            var basis = 0;
            if (arguments.Count() >= 5)
            {
                basis = ArgToInt(arguments, 4);
            }
            if (basis < 0 || basis > 4) return CreateResult(eErrorType.Num);
            var result = IntRateImpl.Intrate(settlementDate, maturityDate, investment, redemption, (DayCountBasis)basis);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, result.DataType);
        }
    }
}
