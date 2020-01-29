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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal class CountBlank : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.First();
            if(!arg.IsExcelRange)throw new InvalidOperationException("CountBlank only support ranges as arguments");
            var result = arg.ValueAsRangeInfo.GetNCells();
            foreach (var cell in arg.ValueAsRangeInfo)
            {
                if (!(cell.Value == null || cell.Value.ToString() == string.Empty))
                {
                    result--;
                }
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
