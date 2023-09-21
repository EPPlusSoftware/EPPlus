using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the discount rate for a security")]
    internal class Disc : ExcelFunction
    {
        public override int ArgumentMinLength => 4;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementNum = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var maturityNum = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            var settlement = DateTime.FromOADate(settlementNum);
            var maturity = DateTime.FromOADate(maturityNum);
            var pr = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if(e3 != null) return CreateResult(e3.Type);
            var redemption = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if(e4 != null) return CreateResult(e4.Type);
            int basis = 0;
            if(arguments.Count > 4)
            {
                basis = ArgToInt(arguments, 4);
            }
            if(maturity <= settlement || pr <= 0 || redemption <= 0 || (basis < 0 || basis > 4))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var yearFrac = new YearFracProvider(context);
            var result = (1d - pr / redemption) / yearFrac.GetYearFrac(settlement, maturity, (DayCountBasis)basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
