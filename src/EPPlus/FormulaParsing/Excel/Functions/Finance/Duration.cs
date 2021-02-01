﻿/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
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
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 5);
            var settlementNum = ArgToDecimal(arguments, 0);
            var maturityNum = ArgToDecimal(arguments, 1);
            var settlement = System.DateTime.FromOADate(settlementNum);
            var maturity = System.DateTime.FromOADate(maturityNum);
            var coupon = ArgToDecimal(arguments, 2);
            var yield = ArgToDecimal(arguments, 3);
            if(coupon < 0 || yield < 0)
            {
                return CreateResult(eErrorType.Num);
            }
            var frequency = ArgToInt(arguments, 4);
            if(frequency != 1 && frequency != 2 && frequency != 4)
            {
                return CreateResult(eErrorType.Num);
            }
            var basis = 0;
            if(arguments.Count() > 5)
            {
                basis = ArgToInt(arguments, 5);
            }
            if(basis < 0 || basis > 4)
            {
                return CreateResult(eErrorType.Num);
            }
            var func = new DurationImpl(new YearFracProvider(context), new CouponProvider());
            var result = func.GetDuration(settlement, maturity, coupon, yield, frequency, (DayCountBasis)basis);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
