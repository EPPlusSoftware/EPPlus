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
using System.Xml.XPath;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of cells (of a supplied range), that satisfy a set of given criteria",
        IntroducedInExcelVersion = "2007")]
    internal class CountIfs : MultipleRangeCriteriasFunction
    {

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var argRanges = new List<RangeOrValue>();
            var criterias = new List<string>();
            for (var ix = 0; ix < 30; ix +=2)
            {
                if (functionArguments.Length <= ix) break;
                var arg = functionArguments[ix];
                if (arg.DataType == DataType.ExcelError) continue;
                var rangeInfo = arg.ValueAsRangeInfo;
                if(rangeInfo == null && arg.ExcelAddressReferenceId > 0)
                {
                    var addressString = ArgToAddress(arguments, ix, context);
                    var address = new ExcelAddress(addressString);
                    var ws = string.IsNullOrEmpty(address.WorkSheetName) ? context.Scopes.Current.Address.Worksheet : address.WorkSheetName;
                    rangeInfo = context.ExcelDataProvider.GetRange(ws, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol, address.Address);
                }
                else if(rangeInfo != null)
                {
                    argRanges.Add(new RangeOrValue { Range = rangeInfo});
                }
                else
                {
                    argRanges.Add(new RangeOrValue { Value = arg.Value });
                }
                var value = functionArguments[ix + 1].Value != null ? ArgToString(arguments, ix + 1) : null;
                criterias.Add(value);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0]);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix]);
                matchIndexes = matchIndexes.Intersect(indexes);
            }
            
            return CreateResult((double)matchIndexes.Count(), DataType.Integer);
        }
    }
}
