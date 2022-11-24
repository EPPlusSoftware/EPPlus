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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns a reference to a cell (or range of cells) for requested rows and columns within a supplied range")]
    internal class Index : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = arguments.ElementAt(0);
            var args = arg1.Value as IEnumerable<FunctionArgument>;
            if (args != null)
            {
                var index = ArgToInt(arguments, 1, RoundingMethod.Floor);
                if (index > args.Count())
                {
                    throw new ExcelErrorValueException(eErrorType.Ref);
                }
                var candidate = args.ElementAt(index - 1);
                //Commented JK-Can be any data type
                //if (!IsNumber(candidate.Value))
                //{
                //    throw new ExcelErrorValueException(eErrorType.Value);
                //}
                //return CreateResult(ConvertUtil.GetValueDouble(candidate.Value), DataType.Decimal);
                return CompileResultFactory.Create(candidate.Value);
            }
            if (arg1.IsExcelRange)
            {
                var row = ArgToInt(arguments, 1, RoundingMethod.Floor);                 
                var col = arguments.Count()>2 ? ArgToInt(arguments, 2, RoundingMethod.Floor) : 1;
                var ri=arg1.ValueAsRangeInfo;
                if (row > ri.Address.ToRow - ri.Address.FromRow + 1 ||
                    col > ri.Address.ToCol - ri.Address.FromCol + 1)
                {
                    return new CompileResult(eErrorType.Ref);
                }
                var candidate = ri.GetOffset(row-1, col-1);
                //Commented JK-Can be any data type
                //if (!IsNumber(candidate.Value))   
                //{
                //    throw new ExcelErrorValueException(eErrorType.Value);
                //}
                return CompileResultFactory.Create(candidate);
            }
            if(arg1.ValueIsExcelError)
            {
                return new CompileResult(arg1.ValueAsExcelErrorValue.Type);
            }
            return new CompileResult(eErrorType.Value);
        }
        public override bool ReturnsReference => true;
    }
}
