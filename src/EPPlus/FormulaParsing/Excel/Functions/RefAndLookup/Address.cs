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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns a reference, in text format, for a supplied row and column number")]
    internal class Address : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var row = ArgToInt(arguments, 0);
            var col = ArgToInt(arguments, 1);
            ThrowExcelErrorValueExceptionIf(() => row < 0 && col < 0, eErrorType.Value);
            var referenceType = ExcelReferenceType.AbsoluteRowAndColumn;
            var worksheetSpec = string.Empty;
            if (arguments.Count() > 2)
            {
                var arg3 = ArgToInt(arguments, 2);
                ThrowExcelErrorValueExceptionIf(() => arg3 < 1 || arg3 > 4, eErrorType.Value);
                referenceType = (ExcelReferenceType)ArgToInt(arguments, 2);
            }
            if (arguments.Count() > 3)
            {
                var fourthArg = arguments.ElementAt(3).Value;
                if (fourthArg is bool && !(bool)fourthArg)
                {
                    throw new InvalidOperationException("Excelformulaparser does not support the R1C1 format!");
                }
            }
            if (arguments.Count() > 4)
            {
                var fifthArg = arguments.ElementAt(4).Value;
                if (fifthArg is string && !string.IsNullOrEmpty(fifthArg.ToString()))
                {
                    worksheetSpec = fifthArg + "!";
                }
            }
            var translator = new IndexToAddressTranslator(context.ExcelDataProvider, referenceType);
            return CreateResult(worksheetSpec + translator.ToAddress(col, row), DataType.ExcelAddress);
        }
    }
}
