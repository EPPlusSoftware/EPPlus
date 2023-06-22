/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/06/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.Financial,
    EPPlusVersion = "7.0",
    Description = "Returns the price of a security having an irregular (long or short) first period. Price is per $100 face value.")]
    internal class Oddfprice : ExcelFunction
    {
        public override int ArgumentMinLength => 8;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0)); //Issue date + 1 (When the security is traded back to the buyer).
            var maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var issueDate = System.DateTime.FromOADate(ArgToInt(arguments, 2));
            var firstCouponDate = System.DateTime.FromOADate(ArgToInt(arguments, 3));
            var rate = ArgToDecimal(arguments, 4); //Interest rate (annually?)
            var yield = ArgToDecimal(arguments, 5); // Security's annual yield
            var redemption = ArgToDecimal(arguments, 6); // The price the company can buy back a security before maturity. This is per $100 FV
            var frequency = ArgToInt(arguments, 7); //Coupon payout frequency. For example: frequency = 2 means semi-annual payouts
            var b = 0;

            if (arguments.Count > 8)
            {
                b = ArgToInt(arguments, 8);

                if (b < 0 || b > 4)
                {
                    return CreateResult(eErrorType.Num);
                }
            }
            var basis = (DayCountBasis)b;

            // Write check to validate that all dates are OK...

            if (rate < 0 || yield <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            // Write check to validate that maturity > first_coupon > settlement > issue
            var daysDefinition = FinancialDaysFactory.Create(basis);

            var sDate = FinancialDayFactory.Create(settlementDate, basis);
            var mDate = FinancialDayFactory.Create(maturityDate, basis);
            var fcDate = FinancialDayFactory.Create(firstCouponDate, basis);
            var iDate = FinancialDayFactory.Create(issueDate, basis);
            
            //var coupDaysFunc = new CoupdaybsImpl(sDate, mDate, frequency, basis);
            //var coupDaysResult = coupDaysFunc.Coupdaybs();
            //var A = coupDaysResult.Result;
            var A = daysDefinition.GetDaysBetweenDates(iDate, sDate);

            //var coupDaysNcFunc = new CoupdaysncImpl(sDate, mDate, frequency, basis);
            //var coupDaysNcResult = coupDaysNcFunc.Coupdaysnc();
            //var DSC = coupDaysNcResult.Result;
            var DSC = daysDefinition.GetDaysBetweenDates(sDate, fcDate);


            //var E = daysDefinition.DaysPerYear / frequency; //Revise
            var coupDaysFunc = new CoupdaysImpl(sDate, fcDate, frequency, basis);
            var coupDaysResult = coupDaysFunc.GetCoupdays();
            var E = coupDaysResult.Result;

            var coupNumFunc = new CoupnumImpl(sDate, mDate, frequency, basis);
            var coupNumResult = coupNumFunc.GetCoupnum();
            var N = coupNumResult.Result; //Revise

            var DFC = daysDefinition.GetDaysBetweenDates(iDate, fcDate);

            if (DFC < E)
            {
                // Short expression

                var t1 = redemption / (System.Math.Pow(yield / frequency + 1, N - 1 + DSC / E));
                var t2 = (100 * rate / frequency * DFC / E) / (System.Math.Pow(1 + yield / frequency, DSC / E));

                var seriet3 = 0d;
                for (var i = 2; i <= N; i++)
                {
                    seriet3 += (100 * rate / frequency) / (System.Math.Pow(1 + yield / frequency, i - 1 + DSC / E));
                }

                var t3 = seriet3;

                var t4 = 100 * rate / frequency * A / E;

                var oddfprice = t1 + t2 + t3 - t4;

                return CreateResult(oddfprice, DataType.Decimal);
            }
            else
            {
                // Long expression

                // Quasi periods: Normal period has to be divided into smaller period that match the frequency
                // The interest in each quasi period is computed and the amounts are summed over the number of quasi
                // coupon periods.

                var coupNumfunc2 = new CoupnumImpl(iDate, fcDate, frequency, basis);
                var coupNumResult2 = coupNumfunc2.GetCoupnum();
                var NC = coupNumResult2.Result;

                // NC number of quasi periods in one odd period

                var quasiNumFunc = new CoupnumImpl(sDate, fcDate, frequency, basis);
                var quasiNumResult = quasiNumFunc.GetCoupnum();
                var Nq = quasiNumResult.Result; // Unsure about this

                var t1 = (redemption) / (System.Math.Pow(1 + yield / frequency, N + Nq + DSC / E));
                var seriet2 = 0d;
                for (var i = 0; i < NC; i++)
                {
                    seriet2 += (i * E / frequency) / (E); // Very unsure about this
                }

                var t2 = (100 * rate / frequency * seriet2) / (System.Math.Pow(1 + yield / frequency, Nq + DSC / E));

                var seriet3 = 0d;
                for (var j = 0; j < N; j++)
                {
                    seriet3 += (100 * rate / frequency) / (System.Math.Pow(1 + yield / frequency, j - Nq + DSC / E));
                }

                var t3 = seriet3;

                var seriet4 = 0d;
                for (var k = 0; k < NC; k++)
                {
                    seriet4 += (E - k * E / frequency) / E; // Vey unsure about this
                }

                var t4 = 100 * rate / frequency * seriet4;

                var oddfprice = t1 + t2 + t3 - t4;

                return CreateResult(oddfprice, DataType.Decimal);


            }
              

        }
    }

}
