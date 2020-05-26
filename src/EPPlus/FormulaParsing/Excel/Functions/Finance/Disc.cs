using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    public class Disc : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            var settlementNum = ArgToDecimal(arguments, 0);
            var maturityNum = ArgToDecimal(arguments, 1);
            var settlement = System.DateTime.FromOADate(settlementNum);
            var maturity = System.DateTime.FromOADate(maturityNum);
            var pr = ArgToDecimal(arguments, 2);
            var redemption = ArgToDecimal(arguments, 3);
            int basis = 0;
            if(arguments.Count() > 4)
            {
                basis = ArgToInt(arguments, 4);
            }
            if(maturity <= settlement || pr <= 0 || redemption <= 0 || (basis < 0 || basis > 4))
            {
                return CreateResult(eErrorType.Num);
            }
            var yearFrac = new YearFracProvider(context);
            var result = (1d - pr / redemption) / yearFrac.GetYearFrac(settlement, maturity, (DayCountBasis)basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
