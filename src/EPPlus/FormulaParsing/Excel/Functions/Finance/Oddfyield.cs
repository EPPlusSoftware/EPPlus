/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

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
        Description = "")]
    internal class Oddfyield : ExcelFunction
    {
        public override int ArgumentMinLength => 8;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0)); //Issue date + 1 (When the security is traded back to the buyer).
            var maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var issueDate = System.DateTime.FromOADate(ArgToInt(arguments, 2)); 
            var firstCouponDate = System.DateTime.FromOADate(ArgToInt(arguments, 3)); 
            var rate = ArgToDecimal(arguments, 4); //Interest rate (annually?)
            var price = ArgToDecimal(arguments, 5); // Security price
            var redemption = ArgToDecimal(arguments, 6); // The price the company can buy back a security before maturity. This is per $100 FV
            var frequency = ArgToInt(arguments, 7); //Coupon payout frequency. For example: frequency = 2 means semi-annual payouts
            var basis = 0;
            if (arguments.Count > 8)
            {
                basis = ArgToInt(arguments, 8);

                if (basis < 0 || basis > 4)
                {
                    return CreateResult(eErrorType.Num);
                }
            }

            // Write check to validate that all dates are OK...

            if (rate < 0 || price <= 0)
            {
                return CreateResult(eErrorType.Num);
            }

            return CreateResult(eErrorType.Name);

            // Write check to validate that maturity > first_coupon > settlement > issue





        }

    }
}
