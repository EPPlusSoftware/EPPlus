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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class LookupFunction : ExcelFunction
    {
        private readonly ValueMatcher _valueMatcher;

        public LookupFunction()
            : this(new LookupValueMatcher())
        {

        }

        public LookupFunction(ValueMatcher valueMatcher)
        {
            _valueMatcher = valueMatcher;
        }

        public override bool IsLookupFuction
        {
            get
            {
                return true;
            }
        }

        protected int IsMatch(object searchedValue, object candidate)
        {
            return _valueMatcher.IsMatch(searchedValue, candidate);
        }

        protected LookupDirection GetLookupDirection(FormulaRangeAddress rangeAddress)
        {
            var nRows = rangeAddress.ToRow - rangeAddress.FromRow;
            var nCols = rangeAddress.ToCol - rangeAddress.FromCol;
            return nCols > nRows ? LookupDirection.Horizontal : LookupDirection.Vertical;
        }

        protected CompileResult Lookup(LookupNavigator navigator, LookupArguments lookupArgs)
        {
            object lastValue = null;
            object lastLookupValue = null;
            int? lastMatchResult = null;
            if (lookupArgs.SearchedValue == null)
            {
                return new CompileResult(eErrorType.NA);
            }
            do
            {
                var matchResult = IsMatch(lookupArgs.SearchedValue, navigator.CurrentValue);
                if (matchResult != 0)
                {
                    if (lastValue != null && navigator.CurrentValue == null) break;

                    if (!lookupArgs.RangeLookup) continue;
                    if (lastValue == null && matchResult > 0)
                    {
                        return new CompileResult(eErrorType.NA);
                    }
                    if (lastValue != null && matchResult > 0 && lastMatchResult < 0)
                    {
                        return CompileResultFactory.Create(lastLookupValue);
                    }
                    lastMatchResult = matchResult;
                    lastValue = navigator.CurrentValue;
                    lastLookupValue = navigator.GetLookupValue();
                }
                else
                {
                    return CompileResultFactory.Create(navigator.GetLookupValue());
                }
            }
            while (navigator.MoveNext());

            return lookupArgs.RangeLookup ? CompileResultFactory.Create(lastLookupValue) : new CompileResult(eErrorType.NA);
        }
    }
}
