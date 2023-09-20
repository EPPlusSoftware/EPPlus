/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.2",
        Description = "Calculates the depreciation of an asset for a specified period, using the double-declining balance method, or some other user-specified method")]
    internal class Ddb : ExcelFunction
    {
        public override int ArgumentMinLength => 4;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var cost = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var salvage = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            var life = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CreateResult(e3.Type);
            var period = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CreateResult(e4.Type);
            var factor = 2d;
            if(arguments.Count >= 5)
            {
                factor = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
                if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            }

            if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0)
                return CreateResult(eErrorType.Num);

            var result = DdbImpl.Ddb(cost, salvage, life, period, factor);
            if (result.HasError) return CreateResult(result.ExcelErrorType);
            return CreateResult(result.Result, DataType.Decimal);
        }
    }
}
