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
        Description = "Calculates the present value of an investment (i.e. the total amount that a series of future periodic constant payments is worth now)")]
    internal class Pv : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var rate = ArgToDecimal(arguments, 0);
            var nPer = ArgToDecimal(arguments, 1);
            var pmt = 0d;
            if (arguments.Count() >= 3)
            {
                pmt = ArgToDecimal(arguments, 2);
            }
            var fv = 0d;
            if (arguments.Count() >= 4)
            {
                fv = ArgToDecimal(arguments, 3);
            }
            var type = 0;
            if (arguments.Count() >= 5)
            {
                type = ArgToInt(arguments, 4);
            }
            var retVal = CashFlowHelper.Pv(rate, nPer, pmt, fv, (PmtDue)type);
            return CreateResult(retVal, DataType.Decimal);
        }
    }
}
