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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Finds the relative position of a value in a supplied array")]
    internal class Match : LookupFunction
    {
        private enum MatchType
        {
            ClosestAbove = -1,
            ExactMatch = 0,
            ClosestBelow = 1
        }

        public Match()
            : base(new WildCardValueMatcher(), new CompileResultFactory())
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);

            var searchedValue = arguments.ElementAt(0).Value;
            var address =  ArgToAddress(arguments,1, context); 
            var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
            var rangeAddress = rangeAddressFactory.Create(address);
            var matchType = GetMatchType(arguments);
            var args = new LookupArguments(searchedValue, address, 0, 0, false, arguments.ElementAt(1).ValueAsRangeInfo);
            var lookupDirection = GetLookupDirection(rangeAddress);
            var navigator = LookupNavigatorFactory.Create(lookupDirection, args, context);
            int? lastValidIndex = null;
            do
            {
                var matchResult = IsMatch(searchedValue, navigator.CurrentValue);

                // For all match types, if the match result indicated equality, return the index (1 based)
                if (searchedValue != null && matchResult == 0)
                {
                    return CreateResult(navigator.Index + 1, DataType.Integer);
                }

                if ((matchType == MatchType.ClosestBelow && matchResult < 0) || (matchType == MatchType.ClosestAbove && matchResult > 0))
                {
                    lastValidIndex = navigator.Index + 1;
                }
                // If matchType is ClosestBelow or ClosestAbove and the match result test failed, no more searching is required
                else if (matchType == MatchType.ClosestBelow || matchType == MatchType.ClosestAbove)
                {
                    break;
                }
            }
            while (navigator.MoveNext());

            if (matchType == MatchType.ExactMatch && !lastValidIndex.HasValue)
                return CreateResult(eErrorType.NA);

            return CreateResult(lastValidIndex, DataType.Integer);
        }

        private MatchType GetMatchType(IEnumerable<FunctionArgument> arguments)
        {
            var matchType = MatchType.ClosestBelow;
            if (arguments.Count() > 2)
            {
                matchType = (MatchType)ArgToInt(arguments, 2);
            }
            return matchType;
        }
    }
}
