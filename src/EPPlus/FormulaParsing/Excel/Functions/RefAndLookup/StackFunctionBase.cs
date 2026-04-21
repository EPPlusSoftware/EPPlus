/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2025          EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class StackFunctionBase : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";

        public override int ArgumentMinLength => 1;

        protected List<IRangeInfo> GetRanges(IEnumerable<FunctionArgument> arguments, out ExcelErrorValue err)
        {
            err = default;
            var ranges = new List<IRangeInfo>();
            foreach (var arg in arguments)
            {
                if (arg.Value is not IRangeInfo)
                {
                    var rng = new InMemoryRange(1, 1);
                    rng.SetValue(0, 0, arg.Value);
                    ranges.Add(rng);
                }
                else
                {
                    var r = arg.ValueAsRangeInfo;
                    if (r == null)
                    {
                        err = ErrorValues.ValueError;
                        break;
                    }
                    ranges.Add(r);
                }
            }
            return ranges;
        }
    }
}
