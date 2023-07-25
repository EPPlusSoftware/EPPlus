using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImConjugate : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            GetComplexNumbers(arguments[0].Value, out double real, out double imag, out string imaginarySuffix);
            if (double.IsNaN(real) || double.IsNaN(imag))
            {
                return CompileResult.GetErrorResult(eErrorType.Num);
            }
            var realPart = real;
            var imagPart = imag * -1;

            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imaginarySuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
