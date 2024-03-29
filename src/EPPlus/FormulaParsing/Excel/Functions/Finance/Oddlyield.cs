﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
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
        EPPlusVersion = "7",
        Description = "Returns yield of a security that has an irregular (odd) last period.")]
    internal class Oddlyield : ExcelFunction
    {
        public override int ArgumentMinLength => 7;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var sd = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var md = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturityDate = DateTime.FromOADate(md);

            var lid = ArgToInt(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            var lastInterestDate = DateTime.FromOADate(lid);
            
            var rate = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            
            var price = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);

            var redemption = ArgToDecimal(arguments, 5, out ExcelErrorValue e6);
            if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
            
            var frequency = ArgToInt(arguments, 6, out ExcelErrorValue e7);
            if (e7 != null) return CompileResult.GetErrorResult(e7.Type);
            
            var b = 0;
            if (arguments.Count > 7)
            {
                b = ArgToInt(arguments, 7, out ExcelErrorValue e8);
                if(e8 != null) return CompileResult.GetErrorResult(e8.Type);
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

            if (!((maturityDate > settlementDate)
                && (settlementDate > lastInterestDate)))
            {
                return CreateResult(eErrorType.Num);
            }

            if (frequency != 1 && frequency != 2 && frequency != 4)
            {
                return CreateResult(eErrorType.Num);
            }

            if (redemption <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            var sDate = FinancialDayFactory.Create(settlementDate, basis);
            var mDate = FinancialDayFactory.Create(maturityDate, basis);
            var liDate = FinancialDayFactory.Create(lastInterestDate, basis);

            var daysDefinition = FinancialDaysFactory.Create(basis);

            var numOfMonths = 12 / frequency;
            var earlyCouponDate = liDate;
            var DCi = 0d;
            var Ai = 0d;
            var dcDivNl = 0d;
            var aDivNl = 0d;
            var DSCDivNl = 0d;

            var coupNumFunc = new CoupnumImpl(earlyCouponDate, mDate, frequency, basis);
            var coupNumResult = coupNumFunc.GetCoupnum();
            var NC = coupNumResult.Result;

            var startDateDatetime = new System.DateTime(1900, 1, 1);
            var endDateDatetime = new System.DateTime(1900, 1, 1);

            var startDate = FinancialDayFactory.Create(startDateDatetime, basis);
            var endDate = FinancialDayFactory.Create(endDateDatetime, basis);

            //Much of the variables below are retrieved from the Microsoft documentation on ODDLYIELD
            //See https://support.microsoft.com/en-us/office/oddlyield-function-c873d088-cf40-435f-8d41-c8232fee9238

            for (var index = 1; index <= NC; index++)
            {
                var lateCouponDate = earlyCouponDate.AddMonths(numOfMonths, earlyCouponDate.Day);
                var NL = daysDefinition.GetDaysBetweenDates(earlyCouponDate, lateCouponDate, true);

                if (index < NC)
                {
                    DCi = NL;
                }
                else
                {
                    DCi = daysDefinition.GetDaysBetweenDates(earlyCouponDate, mDate, true);

                }

                if (lateCouponDate < sDate)
                {
                    Ai = DCi;
                }
                else if (earlyCouponDate < sDate)
                {
                    Ai = daysDefinition.GetDaysBetweenDates(earlyCouponDate, sDate, true);

                    if (Ai < 0)
                    {
                        Ai = 0d;
                    }
                }
                else
                {
                    Ai = 0d;
                }

                if (sDate > earlyCouponDate)
                {
                    startDate = sDate;
                }
                else
                {
                    startDate = earlyCouponDate;
                }

                if (mDate < lateCouponDate)
                {
                    endDate = mDate;
                }
                else
                {
                    endDate = lateCouponDate;
                }

                var DSC = daysDefinition.GetDaysBetweenDates(startDate, endDate, true);

                earlyCouponDate = lateCouponDate;

                dcDivNl += DCi / NL;
                aDivNl += Ai / NL;
                DSCDivNl += DSC / NL;

            }

            //t1, t2, t3, t4 represents different terms for the function that calculates the yield.
            //See https://support.microsoft.com/en-us/office/oddlyield-function-c873d088-cf40-435f-8d41-c8232fee9238

            var t1 = redemption + dcDivNl * 100 * rate / frequency;
            var t2 = price + aDivNl * 100 * rate / frequency;
            var t3 = price + aDivNl * 100 * rate / frequency;
            var t4 = frequency / DSCDivNl;

            var oddlyield = (t1 - t2) / t3 * t4;
            return CreateResult(oddlyield, DataType.Decimal);

        }

    }
}
