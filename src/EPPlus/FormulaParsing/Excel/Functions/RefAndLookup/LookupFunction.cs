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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class LookupFunction : ExcelFunction
    {
        private readonly ValueMatcher _valueMatcher;
        private readonly CompileResultFactory _compileResultFactory;

        public LookupFunction()
            : this(new LookupValueMatcher(), new CompileResultFactory())
        {

        }

        public LookupFunction(ValueMatcher valueMatcher, CompileResultFactory compileResultFactory)
        {
            _valueMatcher = valueMatcher;
            _compileResultFactory = compileResultFactory;
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

        protected LookupDirection GetLookupDirection(RangeAddress rangeAddress)
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
                        return _compileResultFactory.Create(lastLookupValue);
                    }
                    lastMatchResult = matchResult;
                    lastValue = navigator.CurrentValue;
                    lastLookupValue = navigator.GetLookupValue();
                }
                else
                {
                    return _compileResultFactory.Create(navigator.GetLookupValue());
                }
            }
            while (navigator.MoveNext());

            return lookupArgs.RangeLookup ? _compileResultFactory.Create(lastLookupValue) : new CompileResult(eErrorType.NA);
        }
    }
}
