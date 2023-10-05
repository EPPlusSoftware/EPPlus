/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  15/08/2023         EPPlus Software AB           EPPlus v7
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
    EPPlusVersion = "7.0",
    Description = "Calculates amount received at maturity for fully invested bond or other security.")]
    internal class Received : ExcelFunction
    {
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var s = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlement = DateTime.FromOADate(s);

            var m = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturity = DateTime.FromOADate(m);
            
            var investments = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);
            
            var discount = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);
            
            var basis = 0d;
            if (arguments.Count() > 4)
            {
                basis = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
                if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            }
            basis = Math.Floor(basis);
            if (investments <= 0 || discount <= 0) return CreateResult(eErrorType.Num);
            if (basis < 0 || basis > 4) return CreateResult(eErrorType.Num);
            if (settlement >= maturity) return CreateResult(eErrorType.Num);
            var b = (DayCountBasis)basis;
            var daysdefinition = FinancialDaysFactory.Create(b);
            var B = daysdefinition.DaysPerYear;
            var DIM = daysdefinition.GetDaysBetweenDates(settlement, maturity);
            var result = investments / (1d - (discount * DIM / B));
            return CreateResult(result, DataType.Decimal);
        }
    }
}
