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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the sum of a supplied list of numbers")]
    internal class SumSubtotal : HiddenValuesHandlingFunction
    {
        public SumSubtotal()
        {
            IgnoreErrors = false;
        }

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            KahanSum retVal = 0d;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    var c = Calculate(arg, context);
                    if (c is ExcelErrorValue e)
                    {
                        return CompileResult.GetErrorResult(e.Type) ;
                    }
                    else
                    {
                        retVal += (double)c;
                    }
                }
            }
            return CreateResult(retVal.Get(), DataType.Decimal);
        }

        
        private object Calculate(FunctionArgument arg, ParsingContext context)
        {
            KahanSum retVal = 0d;
            if (arg.Value is IRangeInfo ri)
            {
                foreach (var c in ri)
                {
                    if (ri.IsInMemoryRange || !ShouldIgnore(c, context))
                    {
                        //CheckForAndHandleExcelError(c);
                        if (c.IsExcelError) return c.Value;
                        retVal += c.ValueDouble;
                    }
                }
            }
            else
            {
                if(arg.Address != null && ShouldIgnore(arg, context))
                {
                    return retVal.Get();
                }
                retVal += ConvertUtil.GetValueDouble(arg.Value, true);
            }
            return retVal.Get();
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
    }
}
