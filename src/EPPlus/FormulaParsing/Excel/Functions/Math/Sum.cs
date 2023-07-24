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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Returns the sum of a supplied list of numbers")]
    internal class Sum : HiddenValuesHandlingFunction
    {
        public Sum()
        {
            IgnoreErrors = false;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var retVal = 0m;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    retVal += Calculate(arg, context);                    
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        
        private decimal Calculate(FunctionArgument arg, ParsingContext context)
        {
            var retVal = 0m;
            if (ShouldIgnore(arg, context))
            {
                return retVal;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    if(!ShouldIgnore(arg, context))
                    {
                        retVal += Calculate(item, context);
                    }
                }
            }
            else if (arg.Value is IRangeInfo)
            {
                foreach (var c in (IRangeInfo)arg.Value)
                {
                    if (ShouldIgnore(c, context) == false)
                    {
                        CheckForAndHandleExcelError(c);
                        retVal += Convert.ToDecimal(c.ValueDouble);
                    }
                }
            }
            else
            {
                CheckForAndHandleExcelError(arg);
                retVal += ConvertUtil.GetValueDecimal(arg.Value, true);
            }
            return retVal;
        }
    }
}
