using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class RangeOffset : ExcelFunction
    {
        public IRangeInfo StartRange{ get; set; }

        public IRangeInfo EndRange { get; set; }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {

            if (StartRange == null || EndRange == null) return CreateResult(eErrorType.Value);
            
            //Build the address from the minimum row and column to the maximum row and column. StartRange and offsetRange are single cells.
            var fromRow = System.Math.Min(StartRange.Address.FromRow, EndRange.Address.FromRow);
            var toRow = System.Math.Max(StartRange.Address.ToRow, EndRange.Address.ToRow);
            var fromCol = System.Math.Min(StartRange.Address.FromCol, EndRange.Address.FromCol);
            var toCol = System.Math.Max(StartRange.Address.ToCol, EndRange.Address.ToCol);
            var rangeAddress = new RangeInfo(StartRange.Worksheet, new ExcelAddressBase(fromRow, fromCol, toRow, toCol));
            return CreateResult(rangeAddress, DataType.Enumerable);
        } 
    }
}
