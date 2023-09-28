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
            
            var nper = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            
            var pv = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CreateResult(e3.Type);
            
            var startPeriod = ArgToInt(arguments, 3, out ExcelErrorValue e4);
            if(e4 != null) return CreateResult(e4.Type);

            var endPeriod = ArgToInt(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CreateResult(e5.Type);
            
            var type = ArgToInt(arguments, 5, out ExcelErrorValue e6);
            if (e6 != null) return CreateResult(e6.Type);
            
            if (type < 0 || type > 1) return CompileResult.GetErrorResult(eErrorType.Value);
            var result = CumipmtImpl.GetCumipmt(rate, nper, pv, startPeriod, endPeriod, (PmtDue)type);
            if (result.HasError) return CompileResult.GetErrorResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
