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
            var s = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlement = DateTime.FromOADate(s);
            
            var m = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturity = DateTime.FromOADate(m);
            
            var rate = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var pr = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            
            var redemption = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            
            var frequency = ArgToInt(arguments, 5, out ExcelErrorValue e6);
            if (e6 != null) return CompileResult.GetErrorResult(e6.Type);

            var basis = DayCountBasis.US_30_360;
            if(arguments.Count > 6)
            {
                var b = ArgToInt(arguments, 6, out ExcelErrorValue e7);
                if(e7 != null) return CompileResult.GetErrorResult(e7.Type);

                basis = (DayCountBasis)b;
            }
            var func = new YieldImpl(new CouponProvider(), new PriceProvider());
            var result = func.GetYield(settlement, maturity, rate, pr, redemption, frequency, basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
