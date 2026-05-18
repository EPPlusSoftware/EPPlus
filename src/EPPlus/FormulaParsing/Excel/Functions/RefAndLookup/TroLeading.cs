using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.RichData.RichValues.Errors;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "8",
        Description = "",
        SupportsArrays = true)]

    internal class TroLeading : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = arguments[0].ValueAsRangeInfo;
            var result = TrimRangeLeading(range, out var error);
            if (error != null)
                return error;
            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private InMemoryRange TrimRangeLeading(IRangeInfo range, out CompileResult error)
        {
            var startingRow = 0;
            var startingcol = 0;
            error = null;

            for (int col = 0; col < range.Size.NumberOfCols; col++)
            {
                bool startingColFound = false;
                for (int row = 0; row < range.Size.NumberOfRows; row++)
                {
                    if (HasValue(range.GetOffset(row, col))) // Vi hittar en column med värde
                    {                        
                        startingColFound = true;   
                        break;                        
                    }
                }
                if (startingColFound)
                    break;
                startingcol++;
            }

            for (int row = 0; row < range.Size.NumberOfRows; row++)
            {
                bool startingRowFound = false;
                for (int col = 0; col< range.Size.NumberOfCols; col++)
                {
                    if (HasValue(range.GetOffset(row, col))) // Vi hittar en rad med värde
                    {
                        startingRowFound = true;
                        break;
                    }
                }
                if (startingRowFound)
                    break;
                startingRow++;
            }

            if(startingRow == range.Size.NumberOfRows && startingcol == range.Size.NumberOfCols) // range is empty
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }

            int numberOfRows = range.Size.NumberOfRows - startingRow;
            int numberOfCols = range.Size.NumberOfCols - startingcol;

            var result = new InMemoryRange(numberOfRows, (short)numberOfCols);
            for (int r = 0; r < numberOfRows; r++)
                for (int c = 0; c < numberOfCols; c++)
                    result.SetValue(r, c, range.GetOffset(startingRow+ r, startingcol+ c));

            return result;
        }

        private bool HasValue(object val)
        {
            if (val == null) return false;
            if (val.GetType().Name == "RowInternal") return false;
            return true;
        }
    }
}
