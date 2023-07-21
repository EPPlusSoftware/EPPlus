using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal class PriceDisc : ExcelFunction
    {
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            var settlement = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturity = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var discount = ArgToDecimal(arguments, 2);
            var redemption = ArgToDecimal(arguments, 3);
            var b = 0d;

            if (arguments.Count > 4) 
            {
                b = ArgToDecimal(arguments, 4);
            }



            var basis = (DayCountBasis)b;

            if (b < 0 || b > 4)
            {
                return CreateResult(eErrorType.Num);
            }

            if (discount <= 0 || redemption <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            if (settlement >= maturity)
            {
                return CreateResult(eErrorType.Num);
            }

            var daysDefinition = FinancialDaysFactory.Create(basis);

            var DSM = daysDefinition.GetDaysBetweenDates(settlement, maturity);

            var B = daysDefinition.DaysPerYear;

            var pricedisc = redemption - discount * redemption * DSM / B;

            return CreateResult(pricedisc, DataType.Decimal);

        }
    }
}
