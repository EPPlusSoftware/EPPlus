/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/06/2024         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Drawing.Style.Fill;
using System.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Engineering,
       EPPlusVersion = "7.2.1",
       Description = "Returns the annual yield for a discounted security.")]
    internal class YieldDisc : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].IsExcelRange || arguments[1].IsExcelRange || arguments[2].IsExcelRange || arguments[3].IsExcelRange)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var s = ArgToInt(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var settlement = DateTime.FromOADate(s);

            var m = ArgToInt(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            var maturity = DateTime.FromOADate(m);

            var price = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var rvalue = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);

            var basis = 0d;
            if (arguments.Count() > 4)
            {
                basis = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
                if (e5 != null) return CompileResult.GetErrorResult(e5.Type);
            }

            if (price <= 0 || rvalue <= 0 || basis<0 ||basis>4 ||settlement>=maturity)
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }

            var args = new List<FunctionArgument>();
            args.Add(new FunctionArgument(settlement));
            args.Add(new FunctionArgument(maturity));
            args.Add(new FunctionArgument(basis));
            var func = context.Configuration.FunctionRepository.GetFunction("yearfrac");
            var yearfrac = System.Math.Abs(func.Execute(args,context).ResultNumeric);
            var result = ((rvalue - price) / (price*yearfrac));

            return CreateResult(result, DataType.Decimal);
        }   
    }
}