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
        public override int ArgumentMinLength => 4;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            // collect input
            var id = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var issueDate = DateTime.FromOADate(id);

            var sd = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlementDate = DateTime.FromOADate(sd);

            var rate = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CreateResult(e3.Type);

            var par = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if(e4 != null) return CreateResult(e4.Type);
            
            var basis = 0;
            if (arguments.Count > 4)
            {
                basis = ArgToInt(arguments, 4, out ExcelErrorValue e5);
                if(e5 != null) return CreateResult(e5.Type);
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
