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
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns the number of columns in a supplied range")]
    internal class Columns : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var r=arguments.ElementAt(0).ValueAsRangeInfo;
            if (r != null)
            {
                return CreateResult(r.Address.ToCol - r.Address.FromCol + 1, DataType.Integer);
            }
            else
            {
                var range = ArgToAddress(arguments, 0, context);
                if (ExcelAddressUtil.IsValidAddress(range))
                {
                    var factory = new RangeAddressFactory(context.ExcelDataProvider, context);
                    var address = factory.Create(range);
                    return CreateResult(address.ToCol - address.FromCol + 1, DataType.Integer);
                }
            }
            throw new ArgumentException("Invalid range supplied");
        }
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            return FunctionParameterInformation.IgnoreAddress;
        }
    }
}
