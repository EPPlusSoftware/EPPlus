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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns one of a list of values, depending on the value of a supplied index number",
        SupportsArrays = true)]
    internal class Choose : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var indexArg = arguments[0];
            if (indexArg.DataType==DataType.ExcelRange)
            {
                var ri = indexArg.ValueAsRangeInfo;
                var rowFrom = ri.Address?.FromRow ?? 0;
                var colFrom = ri.Address?.FromCol ?? 0;
                var values = new InMemoryRange(ri.Size);
                for (int r = 0; r < ri.Size.NumberOfRows; r++)
                {
                    for (int c = 0; c < ri.Size.NumberOfCols; c++)
                    {
                        var ix = ConvertUtil.ParseInt(ri.GetValue(rowFrom + r, colFrom + c), RoundingMethod.Convert);
                        if(ix<0 && ix >= arguments.Count)
                        {
                            values.SetValue(r, c, ErrorValues.ValueError);
                        }
                        else
                        {
                            values.SetValue(r, c, arguments[ix].Value);
                        }
                    }
                }

                return CreateResult(values, DataType.ExcelRange);
            }
            else
            {
                var index = ArgToInt(arguments, 0, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                var choosedValue = arguments[index].Value;
                if(choosedValue is IRangeInfo ri)
                {
                    return CreateAddressResult(ri, DataType.ExcelRange);
                }
                return CompileResultFactory.Create(choosedValue);
            }
        }
        public override bool ReturnsReference => true;
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
