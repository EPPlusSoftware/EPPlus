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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns the column number of a supplied range, or of the current cell",
        SupportsArrays = true)]
    internal class Column : LookupFunction
    {
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count == 0)
            {
                return CreateResult(context.CurrentCell.Column, DataType.Integer);
            }
            var arg1 = arguments[0];
            if(arg1.IsExcelRange)
            {
                var range = arg1.ValueAsRangeInfo;
                if (range.IsInMemoryRange) return CreateResult(eErrorType.Value);
                if (range.Size.NumberOfCols > 1)
                {
                    var rangeDef = new RangeDefinition(1, range.Size.NumberOfCols);
                    var returnRange = new InMemoryRange(rangeDef);
                    var returnRangeCol = 0;
                    for (var col = range.Address.FromCol; col <= range.Address.ToCol; col++)
                    {
                        returnRange.SetValue(0, returnRangeCol++, col);
                    }
                    return CreateResult(returnRange, DataType.ExcelRange);
                }
                else
                {
                    return CreateResult(range.Address.FromCol, DataType.Integer);
                }
            }
            else
            {
                var rangeAddress = ArgToAddress(arguments, 0);
                if (!ExcelAddressUtil.IsValidAddress(rangeAddress))
                    throw new ArgumentException("An invalid argument was supplied");
                var factory = new RangeAddressFactory(context.ExcelDataProvider, context);
                var address = factory.Create(rangeAddress);
                return CreateResult(address.FromCol, DataType.Integer);
            }
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreAddress;
        }));
        public override int ArgumentMinLength => 0;
        public override bool IsVolatile => true; //Blank argument will return the current cells column, so set volatile
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
