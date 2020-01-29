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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Exact : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var val1 = arguments.ElementAt(0).ValueFirst;
            var val2 = arguments.ElementAt(1).ValueFirst;

            if (val1 == null && val2 == null)
            {
                return CreateResult(true, DataType.Boolean);
            }
            else if ((val1 == null && val2 != null) || (val1 != null && val2 == null))
            {
                return CreateResult(false, DataType.Boolean);
            }

            var result = string.Compare(val1.ToString(), val2.ToString(), StringComparison.Ordinal);
            return CreateResult(result == 0, DataType.Boolean);
        }
    }
}
