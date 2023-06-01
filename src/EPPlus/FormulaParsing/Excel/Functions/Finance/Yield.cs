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
       Description = "Calculates the yield of a security that pays periodic interest")]
    internal class Yield : ExcelFunction
    {
        public override int ArgumentMinLength => 6;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlement = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var maturity = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var rate = ArgToDecimal(arguments, 2);
            var pr = ArgToDecimal(arguments, 3);
            var redemption = ArgToDecimal(arguments, 4);
            var frequency = ArgToInt(arguments, 5);
            var basis = DayCountBasis.US_30_360;
            if(arguments.Count > 6)
            {
                basis = (DayCountBasis)ArgToInt(arguments, 6);
            }
            var func = new YieldImpl(new CouponProvider(), new PriceProvider());
            var result = func.GetYield(settlement, maturity, rate, pr, redemption, frequency, basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
