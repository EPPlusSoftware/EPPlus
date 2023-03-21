/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
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
        EPPlusVersion = "6.0",
        Description = "Calculates he accrued interest for a security that pays interest at maturity.")]
    internal class AccrintM : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            // collect input
            var issueDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            var rate = ArgToDecimal(arguments, 2);
            var par = ArgToDecimal(arguments, 3);
            var basis = 0;
            if (arguments.Count() > 4)
            {
                basis = ArgToInt(arguments, 4);
            }

            if (rate <= 0 || par <= 0) return CreateResult(eErrorType.Num);
            if (basis < 0 || basis > 4) return CreateResult(eErrorType.Num);
            if (issueDate >= settlementDate) return CreateResult(eErrorType.Num);

            var dayCountBasis = (DayCountBasis)basis;
            var fd = FinancialDaysFactory.Create(dayCountBasis);
            var result = fd.GetDaysBetweenDates(issueDate, settlementDate)/fd.DaysPerYear * rate * par;
            return CreateResult(result, DataType.Decimal);

        }
    }
}
