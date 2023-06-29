using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal class ImProduct : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var realPart = 0d;
            var imagPart = 0d;
            var imSuffix = string.Empty;

            foreach(var arg in arguments)
            {
                GetComplexNumbers(arg.Value.ToString(), out double real, out double imag, out string imaginarySuffix);
                if (double.IsNaN(real) || double.IsNaN(imag))
                {
                    return CompileResult.GetErrorResult(eErrorType.Num);
                }
                if (!string.IsNullOrEmpty(imaginarySuffix))
                {
                    if (!string.IsNullOrEmpty (imSuffix) && imSuffix != imaginarySuffix) 
                    {
                        return CompileResult.GetErrorResult(eErrorType.Value);
                    }
                    imSuffix = imaginarySuffix;
                }
                realPart *= real;
                imagPart *= imag;
            }
            if (imagPart == 0)
            {
                return CreateResult(realPart.ToString(), DataType.String);
            }
            var sign = (imagPart < 0) ? "-" : "+";
            var result = CreateImaginaryString(realPart, imagPart, sign, imSuffix);
            return CreateResult(result, DataType.String);
        }
    }
}
