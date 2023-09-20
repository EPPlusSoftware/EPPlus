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
        Description = "Calculates the cumulative interest paid between two specified periods")]
    internal class Cumipmt : ExcelFunction
    {
        public override int ArgumentMinLength => 6;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rate = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var nper = ArgToInt(arguments, 1);
            var pv = ArgToDecimal(arguments, 2, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            var startPeriod = ArgToInt(arguments, 3);
            var endPeriod = ArgToInt(arguments, 4);
            var type = ArgToInt(arguments, 5);
            if (type < 0 || type > 1) return CompileResult.GetErrorResult(eErrorType.Value);
            var result = CumipmtImpl.GetCumipmt(rate, nper, pv, startPeriod, endPeriod, (PmtDue)type);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
