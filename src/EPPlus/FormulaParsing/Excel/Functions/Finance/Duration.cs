/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
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
        Description = "Calculates the Macauley duration of a security with an assumed par value of $100")]
    internal class Duration : ExcelFunction
    {
        public override int ArgumentMinLength => 5;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementNum = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            
            var maturityNum = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if(e2 != null) return CreateResult(e2.Type);
            
            var settlement = DateTime.FromOADate(settlementNum);
            var maturity = DateTime.FromOADate(maturityNum);
            
            var coupon = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if(e3 != null) return CreateResult(e3.Type);
            
            var yield = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CreateResult(e4.Type);
            
            if(coupon < 0 || yield < 0)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var frequency = ArgToInt(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            if(frequency != 1 && frequency != 2 && frequency != 4)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var basis = 0;
            if(arguments.Count > 5)
            {
                basis = ArgToInt(arguments, 5, out ExcelErrorValue e6);
                if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
            }
            if(basis < 0 || basis > 4)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var func = new DurationImpl(new YearFracProvider(context), new CouponProvider());
            var result = func.GetDuration(settlement, maturity, coupon, yield, frequency, (DayCountBasis)basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
