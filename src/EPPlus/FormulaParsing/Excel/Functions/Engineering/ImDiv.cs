using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImDiv : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = ArgToString(arguments, 0);
            var arg2 = ArgToString(arguments, 1);

            GetComplexNumbers(arg1, out double real, out double imag, out string imaginarySuffix);
            GetComplexNumbers(arg2, out double real2, out double imag2, out string imaginarySuffix2);

            var prelRealPart = (real * real2) + (imag * imag2);
            var prelImagPart = (imag * real2) - (real * imag2);
            var toDivideWith = (real2 * real2) + (imag2 * imag2);
            var realPart = (prelRealPart / toDivideWith);
            var imagPart = (prelImagPart / toDivideWith);
            var sign = (imagPart < 0) ? "-" : "+";

            var usedPrefixes = GetUniquePrefixes(imaginarySuffix, imaginarySuffix2);
            var imSuffix = string.Empty;
            if (usedPrefixes.Count > 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            else if (usedPrefixes.Count == 1)
            {
                imSuffix = usedPrefixes[0];
            }
            var result = CreateImaginaryString(realPart, imagPart, sign, imSuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
