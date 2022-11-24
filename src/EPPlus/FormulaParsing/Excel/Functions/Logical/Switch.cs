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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "5.0",
        Description = "Compares a number of supplied values to a supplied test expression and returns a result corresponding to the first value that matches the test expression. ",
        IntroducedInExcelVersion = "2019")]
    internal class Switch : ExcelFunction
    {
        public Switch()
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var expression = arguments.ElementAt(0).ValueFirst;
            var maxLength = 1 + 126 * 2;
            for(var x = 1; x < (arguments.Count() - 1) || x >= maxLength; x += 2)
            {
                var candidate = arguments.ElementAt(x).Value;
                if(IsMatch(expression, candidate))
                {
                    return CompileResultFactory.Create(arguments.ElementAt(x + 1).Value);
                }
            }
            if (arguments.Count() % 2 == 0) return CompileResultFactory.Create(arguments.Last().Value);
            return new CompileResult(eErrorType.NA);
        }

        private bool IsMatch(object right, object left)
        {
            if(IsNumeric(right) || IsNumeric(left))
            {
                var r = GetNumericValue(right);
                var l = GetNumericValue(left);
                return r.Equals(l);
            }
            if(right == null && left == null)
            {
                return true;
            }
            if (right == null) return false;
            return right.Equals(left);
        }

        private double GetNumericValue(object obj)
        {
            if(obj is System.DateTime)
            {
                return ((System.DateTime)obj).ToOADate();
            }
            if(obj is TimeSpan)
            {
                return ((TimeSpan)obj).TotalMilliseconds;
            }
            return Convert.ToDouble(obj);
        }
        public override bool ReturnsReference => true;
    }
}
