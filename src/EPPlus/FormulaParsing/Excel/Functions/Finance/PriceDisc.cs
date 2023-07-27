/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2023         EPPlus Software AB       Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
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
   EPPlusVersion = "7.0",
   Description = "Calculates the price per 100$ Face value of a discounted bond or other security.")]
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
                b = Math.Floor(b);
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
