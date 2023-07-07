/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/07/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the sum of a supplied list of numbers")]
    internal class SumV2 : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var retVal = 0d;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    var c = Calculate(arg, context, out eErrorType? errType);
                    if (errType.HasValue)
                    {
                        return CompileResult.GetErrorResult(errType.Value);
                    }
                    else if(!double.IsNaN(c))
                    {
                        retVal += c;
                    }
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double Calculate(FunctionArgument arg, ParsingContext context, out eErrorType? errType)
        {
            var retVal = 0d;
            errType = default;
            if (arg.DataType == DataType.ExcelError)
            {
                errType = arg.ValueAsExcelErrorValue.Type;
                return double.NaN;
            }
            if (arg.Value is IRangeInfo ri)
            {
                foreach (var c in ri)
                {
                    if (c.IsExcelError)
                    {
                        errType = ((ExcelErrorValue)c.Value).Type;
                        return double.NaN;
                    }
                    retVal += c.ValueDouble;
                }
            }
            else if(arg.DataType == DataType.ExcelError)
            {
                errType = arg.ValueAsExcelErrorValue.Type;
                return double.NaN;
            }
            else if (arg.DataType == DataType.Boolean && arg.Address != null)
            {
                return 0d;
            }
            else
            {  
                retVal += ConvertUtil.GetValueDouble(arg.Value);
            }
            return retVal;
        }
    }
}
