/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  26/06/2023         EPPlus Software AB           EPPlus v7
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
using System.Runtime.InteropServices;
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
            var firstCouponDate = DateTime.FromOADate(fcd);

            var rate = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);

            var yield = ArgToDecimal(arguments, 5, out ExcelErrorValue e6);
            if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
            
            var redemption = ArgToDecimal(arguments, 6, out ExcelErrorValue e7);
            if(e7 != null) return CompileResult.GetErrorResult(e7.Type);

            var frequency = ArgToInt(arguments, 7, out ExcelErrorValue e8);
            if (e8 != null) return CompileResult.GetErrorResult(e8.Type);
            
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

            if (rate < 0 || yield < 0)
            {
                return CreateResult(eErrorType.Num);
            }

            if (b < 0 || b > 4)
            {
                return CreateResult(eErrorType.Num);
            }

            var oddfpriceFunc = new OddfpriceImpl(settlementDate, maturityDate, issueDate, firstCouponDate, rate, yield, redemption, frequency, basis);
            var result = oddfpriceFunc.GetOddfprice();
            if (result.HasError)
            {
                return CreateResult(result.ExcelErrorType);
            }


            return CreateResult(result.Result, DataType.Decimal);
              

        }
    }

}
