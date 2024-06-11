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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns the number of rows in a supplied range")]
    internal class Rows : LookupFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var r = arguments[0].ValueAsRangeInfo;
            if (r != null)
            {
                return CreateResult(r.Address.ToRow - r.Address.FromRow + 1, DataType.Integer);
            }
            else
            {
                var range = ArgToAddress(arguments, 0);
                if (ExcelAddressUtil.IsValidAddress(range))
                {
                    var factory = new RangeAddressFactory(context.ExcelDataProvider, context);
                    var address = factory.Create(range);
                    return CreateResult(address.ToRow - address.FromRow + 1, DataType.Integer);
                }
            }
            if(context.Debug)
            {
                context.Configuration.Logger.Log("Rows function:Invalid range supplied. Cell {context.CurrentWorksheet?.Name}!{context.CurrentCell?.Address}");
            }
            return CompileResult.GetErrorResult(eErrorType.Value);
        }
        /// <summary>
        /// Reference Parameters do not need to be follows in the dependency chain.
        /// </summary>
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreAddress;
        }));
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
