/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
         Category = ExcelFunctionCategory.Statistical,
         EPPlusVersion = "5.5",
         Description = "Returns the variance of a supplied set of values (which represent a sample of a population), counting text and the logical value FALSE as the value 0 and counting the logical value TRUE as the value 1")]
    internal class Vara : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments.Any() || arguments.Count() < 2) return CreateResult(eErrorType.Div0);
            var varFunc = new Var();
            var values = new List<double>();
            GetValues(arguments, values);
            var newArgs = values.Select(x => new FunctionArgument(x));
            var result = varFunc.Execute(newArgs, context);
            return result;
        }

        private void GetValues(IEnumerable<FunctionArgument> arguments, List<double> values)
        {
            foreach(var arg in arguments)
            {
                if(arg.IsExcelRange)
                {
                    foreach(var cell in arg.ValueAsRangeInfo)
                    {
                        HandleValue(cell.Value, values);
                    }
                }
                else
                {
                    HandleValue(arg.Value, values);
                }
            }
        }

        private void HandleValue(object val, List<double> values)
        {
            if (val == null) return;
            if (IsNumeric(val))
            {
                values.Add(ConvertUtil.GetValueDouble(val));
            }
            else if(!string.IsNullOrEmpty(val.ToString()))
            {
                values.Add(0d);
            }
        }
    }
}
