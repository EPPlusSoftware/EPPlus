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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System.Text.RegularExpressions;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Searches for a specific value in one data vector, and returns a value from the corresponding position of a second data vector")]
    internal class Lookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            if (HaveTwoRanges(arguments))
            {
                return HandleTwoRanges(arguments, context);
            }
            return HandleSingleRange(arguments, context);
        }

        private bool HaveTwoRanges(IEnumerable<FunctionArgument> arguments)
        {
            if (arguments.Count() < 3) return false;
            return (arguments.ElementAt(2).Value is RangeInfo);
        }

        private CompileResult HandleSingleRange(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments.ElementAt(0).Value;
            Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
            var firstAddress = ArgToAddress(arguments, 1, context);
            var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider, context);
            var address = rangeAddressFactory.Create(firstAddress);
            var nRows = address.ToRow - address.FromRow;
            var nCols = address.ToCol - address.FromCol;
            var lookupIndex = nCols + 1;
            var lookupDirection = LookupDirection.Vertical;
            if (nCols > nRows)
            {
                lookupIndex = nRows + 1;
                lookupDirection = LookupDirection.Horizontal;
            }
            var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, 0, true, arguments.ElementAt(1).ValueAsRangeInfo);
            var navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);
            return Lookup(navigator, lookupArgs);
        }

        private CompileResult HandleTwoRanges(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var searchedValue = arguments.ElementAt(0).Value;
            Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
            Require.That(arguments.ElementAt(2).Value).Named("secondAddress").IsNotNull();
            var firstAddress = ArgToAddress(arguments, 1, context);
            var secondAddress = ArgToAddress(arguments, 2, context);
            var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider, context);
            var address1 = rangeAddressFactory.Create(firstAddress);
            var address2 = rangeAddressFactory.Create(secondAddress);
            var lookupIndex = (address2.FromCol - address1.FromCol) + 1;
            var lookupOffset = address2.FromRow - address1.FromRow;
            var lookupDirection = GetLookupDirection(address1);
            if (lookupDirection == LookupDirection.Horizontal)
            {
                lookupIndex = (address2.FromRow - address1.FromRow) + 1;
                lookupOffset = address2.FromCol - address1.FromCol;
            }
            var lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, lookupOffset,  true, arguments.ElementAt(1).ValueAsRangeInfo);
            var navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);
            return Lookup(navigator, lookupArgs);
        }
    }
}
