/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/06/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "7.0",
        Description = "Returns the yield of a bond or other security that has a long or short first period.")]
    internal class Oddfyield : ExcelFunction
    {
        public override int ArgumentMinLength => 8;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var sd = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var md = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturityDate = DateTime.FromOADate(md);

            var id = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            var issueDate = DateTime.FromOADate(id);

            var fcd = ArgToInt(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            var firstCouponDate = System.DateTime.FromOADate(fcd);
            
            var rate = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            
            var price = ArgToDecimal(arguments, 5, out ExcelErrorValue e6);
            if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
            
            var redemption = ArgToDecimal(arguments, 6, out ExcelErrorValue e7);
            if (e7 != null) return CompileResult.GetErrorResult(e7.Type);

            var frequency = ArgToInt(arguments, 7, out ExcelErrorValue e8);
            if(e8 != null) return CompileResult.GetErrorResult(e8.Type);

            var b = 0;
            if (arguments.Count > 8)
            {
                b = ArgToInt(arguments, 8, out ExcelErrorValue e9);
                if (e9 != null) return CompileResult.GetErrorResult(e9.Type);

                if (b < 0 || b > 4)
                {
                    return CreateResult(eErrorType.Num);
                }
            }

            var basis = (DayCountBasis)b;


            if (rate < 0 || price <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            var daysDefinition = FinancialDaysFactory.Create(basis);

            //Excel uses Newton-Raphson method to calculate ODDFYIELD.
            //It uses ODDFPRICE by tweaking the yield parameter to get ODDFPRICE to match the price argument. See implementation below

            var sDate = FinancialDayFactory.Create(settlementDate, basis);
            var mDate = FinancialDayFactory.Create(maturityDate, basis);

            //Excel uses Newton-Raphson method to calculate ODDFYIELD.
            //It uses ODDFPRICE by tweaking the yield parameter to get ODDFPRICE to match the price argument. See implementation below

            var daysTilMaturity = daysDefinition.GetDaysBetweenDates(sDate, mDate);
            var epsilon = 0.00000001d;
            var numerator = rate * daysTilMaturity * 100 - (price - 100);
            var denominator = (price - 100) / 4 + daysTilMaturity * (price - 100) / 2 + daysTilMaturity * 100;
            var guess = numerator / denominator;
            var estimatedYield = 0d; //We know the price, so lets tweak this parameter when doing the ODDFPRICE-calculations, and since we know the price, we can have a tolerable margin of error.

            //Now we want to simulate with 100 max iterations bond prices through newton-rhapson method. Satisfactory when error < epsilon (1 * 10^(-8))

            for (var iteration = 0; iteration <= 100; iteration++)
            {
                estimatedYield = guess / frequency;
                var bondPriceEstimationFunc = new OddfpriceImpl(settlementDate, maturityDate, issueDate, firstCouponDate, rate, estimatedYield, redemption, frequency, basis);
                var bondPriceEstimation = bondPriceEstimationFunc.GetOddfprice();

                if (bondPriceEstimation.HasError)
                {
                    return CreateResult(bondPriceEstimation.ExcelErrorType);
                }


                var estimatedError = bondPriceEstimation.Result - price;

                if (System.Math.Abs(estimatedError) <= epsilon )
                {
                    return CreateResult(estimatedYield, DataType.Decimal);
                }

                var yieldEpsilon = (guess + epsilon) / frequency;
                var bondPriceEpsilonFunc = new OddfpriceImpl(settlementDate, maturityDate, issueDate, firstCouponDate, rate, yieldEpsilon, redemption, frequency, basis);
                var bondPriceEpsilon = bondPriceEpsilonFunc.GetOddfprice();
                if (bondPriceEpsilon.HasError)
                {
                    return CreateResult(bondPriceEpsilon.ExcelErrorType);
                }

                var slope = (bondPriceEpsilon.Result - bondPriceEstimation.Result) / epsilon; // The derivative of the estimated bond price

                guess -= estimatedError / slope;

                iteration += 1;

            }

            return CreateResult(estimatedYield, DataType.Decimal);

        }

    }
}
