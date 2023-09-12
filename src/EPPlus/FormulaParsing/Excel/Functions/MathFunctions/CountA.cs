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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of non-blanks in a supplied set of cells or values")]
    internal class CountA : HiddenValuesHandlingFunction
    {
        public CountA() : base()
        {
            IgnoreErrors = false;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if(argumentIndex == 0) { }
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var nItems = 0d;
            Calculate(arguments, context, ref nItems, out ExcelErrorValue e);
            if(e != null)
            {
                return CreateResult(e.Type);
            }
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ParsingContext context, ref double nItems, out ExcelErrorValue err)
        {
            err = default;
            foreach (var item in items)
            {
                var cs = item.Value as IRangeInfo;
                if (cs != null)
                {
                    foreach (var c in cs)
                    {
                        if (!ShouldIgnore(c, context) && ShouldCount(c.Value))
                        {
                            nItems++;
                        }
                    }
                }
                else if (item.Value is IEnumerable<FunctionArgument>)
                {
                    Calculate((IEnumerable<FunctionArgument>)item.Value, context, ref nItems, out ExcelErrorValue e);
                    if(e != null)
                    {
                        err = e;
                        return;
                    }
                }
                else
                {
                    if (!ShouldIgnore(item, context) && ShouldCount(item.Value))
                    {
                        nItems++;
                    }
                }
            }
        }

        private bool ShouldCount(object value)
        {
            return value != null;
        }
    }
}
