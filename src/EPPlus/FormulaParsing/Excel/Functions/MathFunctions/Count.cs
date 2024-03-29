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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of numerical values in a supplied set of cells or values")]
    internal class Count : HiddenValuesHandlingFunction
    {
        private enum ItemContext
        {
            InRange,
            SingleArg
        }

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var nItems = 0d;
            Calculate(arguments, ref nItems, context, ItemContext.SingleArg, out eErrorType? error);
            if (error.HasValue && IgnoreErrors==false)
            {
                return CompileResult.GetErrorResult(error.Value);
            }
            else
            {
                return CreateResult(nItems, DataType.Integer);
            }
        }

        private void Calculate(IList<FunctionArgument> items, ref double nItems, ParsingContext context, ItemContext itemContext, out eErrorType? error)
        {
            error = null;
            foreach (var item in items)
            {
                var cs = item.Value as IRangeInfo;
                if (cs != null)
                {
                    foreach (var c in cs)
                    {
                        if(c is ExcelErrorValue ev)
                        {
                            error = ev.Type;
                            return;
                        }
                        if (ShouldIgnore(c, context) == false && ShouldCount(c.Value, ItemContext.InRange))
                        {
                            nItems++;
                        }
                    }
                }
                else
                {
                    if (item.DataType == DataType.ExcelError)
                    {
                        error = item.ValueAsExcelErrorValue.Type;
                    }
                    if (ShouldIgnore(item, context) == false && ShouldCount(item.Value, itemContext))
                    {
                        nItems++;
                    }
                }
            }
        }

        private bool ShouldCount(object value, ItemContext context)
        {
            switch (context)
            {
                case ItemContext.SingleArg:
                    return IsNumeric(value) || IsNumericString(value);
                case ItemContext.InRange:
                    return ConvertUtil.IsExcelNumeric(value);
                default:
                    throw new ArgumentException("Unknown ItemContext:" + context.ToString());
            }
        }
    }
}
