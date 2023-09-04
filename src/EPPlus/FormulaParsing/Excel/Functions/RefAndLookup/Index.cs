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
            var crf = new CompileResultFactory();
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
                return crf.Create(candidate.Value);
            }
            if (arg1.IsExcelRange)
            {
                var row = ArgToInt(arguments, 1, RoundingMethod.Floor);                 
                var col = arguments.Count()>2 ? ArgToInt(arguments, 2, RoundingMethod.Floor) : 0;
                var ri=arg1.ValueAsRangeInfo;
                var nRows = ri.Address._toRow - ri.Address._fromRow + 1;
                var nCols = ri.Address._toCol - ri.Address._fromCol + 1;
                var rowIx = row - 1;
                var colIx = col - 1;
                if(nRows == 1 && row <= nCols && col == 0)
                {
                    colIx = rowIx;
                    rowIx = 0;
                }
                else if (row >  nRows ||  col > nCols || (col == 0 && nRows > 1 && nCols > 1))
                {
                    return CreateResult(eErrorType.Ref);
                }
                var candidate = ri.GetOffset(rowIx, colIx < 0 ? 0 : colIx);
                //Commented JK-Can be any data type
                //if (!IsNumber(candidate.Value))   
                //{
                //    throw new ExcelErrorValueException(eErrorType.Value);
                //}
                return crf.Create(candidate);
            }
            else
            {
                // only one argument
                if(arg1.ValueIsExcelError)
                {
                    return CreateResult(arg1.Value, DataType.ExcelError);
                }
                else if (arguments.ElementAt(1).ValueIsExcelError)
                {
                    return CreateResult(arguments.ElementAt(1).Value, DataType.ExcelError);
                }
                var index = ArgToInt(arguments, 1, RoundingMethod.Floor);
                return index >= 0 && index <= 1 ? crf.Create(arg1.Value) : CreateResult(eErrorType.Ref);
            }
        }
    }
}
