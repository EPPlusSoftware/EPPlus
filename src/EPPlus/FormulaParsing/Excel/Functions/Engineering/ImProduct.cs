using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations;
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
            var args = new List <string> ();
            if (arguments[0].IsExcelRange)
            {
                var range = arguments[0].ValueAsRangeInfo;
                foreach (var cell in range)
                {
                    var cellValue = cell.Value;
                    if (cellValue != null)
                    {
                        args.Add(cellValue.ToString());
                    }
                }
            }
            else
            {
                foreach (var arg in arguments)
                {
                    args.Add(arg.Value.ToString());
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
            #region old code
            //var realPart1 = 1d;
            //var imagPart1 = 1d;
            //var realPart2 = 1d;
            //var imagPart2 = 1d;


            //double x = 0;
            //double y = 0;
            //double u = 0;
            //double v = 0;

            //int index = 0;

            ////var lastReal = 1d;
            ////var lastImag = 1d;

            //for(int i = 0; i < 2; i++)
            //{
            //    GetComplexNumbers(arguments[i].Value.ToString(), out double real, out double imag, out string imaginarySuffix);
            //    if (double.IsNaN(real) || double.IsNaN(imag))
            //    {
            //        return CompileResult.GetErrorResult(eErrorType.Num);
            //    }
            //    if (!string.IsNullOrEmpty(imaginarySuffix))
            //    {
            //        if (!string.IsNullOrEmpty(imSuffix) && imSuffix != imaginarySuffix)
            //        {
            //            return CompileResult.GetErrorResult(eErrorType.Value);
            //        }
            //        imSuffix = imaginarySuffix;
            //    }

            //    if(x == 0 && y == 0)
            //    {
            //        x = real; 
            //        y = imag;
            //    }
            //    else
            //    {
            //        u = real; 
            //        v = imag;
            //    }
            //}

            //var realResult = x * u;
            //var imagResult = y * v;

            //foreach (var arg in arguments)
            //{
            //    GetComplexNumbers(arg.Value.ToString(), out double real, out double imag, out string imaginarySuffix);
            //    if (double.IsNaN(real) || double.IsNaN(imag))
            //    {
            //        return CompileResult.GetErrorResult(eErrorType.Num);
            //    }
            //    if (!string.IsNullOrEmpty(imaginarySuffix))
            //    {
            //        if (!string.IsNullOrEmpty (imSuffix) && imSuffix != imaginarySuffix) 
            //        {
            //            return CompileResult.GetErrorResult(eErrorType.Value);
            //        }
            //        imSuffix = imaginarySuffix;
            //    }

            //    //if(!isFirst)
            //    //{
            //    //    x = real;
            //    //    y = imag;
            //    //    isFirst = true;
            //    //}
            //    //else
            //    //{
            //    //    u = real; v = imag;
            //    //}

            //    //realPart = ((real*real) - (imag*imag));
            //    //imagPart = ((real*imag) + (real*imag));

            //    imagPart1 = (realPart1 * real);
            //    imagPart2 = (realPart2 * imag);

            //    realPart1 *= imag;
            //    realPart2 *= real;

            //    index++;

            //    if((index % 2) == 0)
            //    {
            //        realPart1 = realPart1 - realPart2;
            //        realPart2 = imagPart1 + imagPart2;
            //    }
            //}

            //var realRes = (x * u) - (y * v);
            //var imRes = (x * v) - (y * u);

            // var res1 = realRes + "+" + imRes;



            ////var realResult = realPart1 - realPart2; 
            ////var imagResult = imagPart1 + imagPart2;

            //if (imagResult == 0)
            //{
            //    return CreateResult(realResult.ToString(), DataType.String);
            //}//var sign = (imagResult < 0) ? "-" : "+";
            //var result = CreateImaginaryString(realResult, imagResult, sign, imSuffix);



            //arguments.RemoveAt(0);
            //arguments.RemoveAt(1);
            //arguments.Insert(0, result);
            //Execute()
            #endregion
            return CreateResult(result, DataType.String);
        }
    }
}
