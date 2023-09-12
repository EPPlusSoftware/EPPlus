using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Engineering,
       EPPlusVersion = "7.0",
       Description = "Returns the product of 1 to 255 complex numbers in x + yi or x + yj text format.")]
    internal class ImProduct : ImFunctionBase
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var args = new List <string> ();
            foreach(var arg in arguments)
            {
                if (arguments[0].IsExcelRange)
                {
                    var range = arguments[0].ValueAsRangeInfo;
                    foreach (var cell in range)
                    {
                        var cellValue = cell.Value;
                        if (cellValue == null)
                        {
                            args.Add("0");
                        }
                        else
                        {
                            args.Add(cellValue.ToString());
                        }
                    }
                }
                else
                {                    
                    if(arg.DataType!=DataType.Empty)
                    {
                        args.Add(arg.Value.ToString());
                    }
                }
            }
            var imSuffix = string.Empty;
            ComplexNumber complexResult = default;
            foreach (var argument in args) 
            {
                GetComplexNumbers(argument, out double real, out double imag, out string imaginarySuffix);
                if (double.IsNaN(real) || double.IsNaN(imag))
                {
                    return CompileResult.GetErrorResult(eErrorType.Num);
                }
                if (!string.IsNullOrEmpty(imaginarySuffix))
                {
                    if (!string.IsNullOrEmpty(imSuffix) && imSuffix != imaginarySuffix)
                    {
                        return CompileResult.GetErrorResult(eErrorType.Value);
                    }
                    imSuffix = imaginarySuffix;
                }

                var complexNumber = new ComplexNumber(real, imag, imSuffix);
                if(complexResult == default)
                {
                    complexResult = complexNumber;
                }
                else
                {
                    complexResult = complexResult.GetProduct(complexNumber);
                }
            }
            var sign = (complexResult.Imaginary < 0) ? "-" : "+";
            var result = CreateImaginaryString(complexResult.Real, complexResult.Imaginary, sign, imSuffix);

            return CreateResult(result, DataType.String);
        }
    }
}
