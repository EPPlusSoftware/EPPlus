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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class ArrayLookupNavigator : LookupNavigator
    {
        private readonly FunctionArgument[] _arrayData;
        private int _index = 0;
        private object _currentValue;
 
        public ArrayLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
            : base(direction, arguments, parsingContext)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            Require.That(arguments.DataArray).Named("arguments.DataArray").IsNotNull();
            _arrayData = arguments.DataArray.ToArray();
            Initialize();
        }

        private void Initialize()
        {
            if (Arguments.LookupIndex >= _arrayData.Length)
            {
                throw new ExcelErrorValueException(eErrorType.Ref);
            }
            SetCurrentValue();

        }

        public override int Index
        {
            get { return _index; }
        }

        private void SetCurrentValue()
        {
            _currentValue = _arrayData[_index];
        }

        private bool HasNext()
        {
            if (Direction == LookupDirection.Vertical)
            {
                return _index < (_arrayData.Length - 1);
            }
            else
            {
                return false;
            }
        }

        public override bool MoveNext()
        {
            if (!HasNext()) return false;
            if (Direction == LookupDirection.Vertical)
            {
                _index++;
            }
            SetCurrentValue();
            return true;
        }

        public override object CurrentValue
        {
            get { return _arrayData[_index].Value; }
        }

        public override object GetLookupValue()
        {
            return _arrayData[_index].Value;
        }
    }
}
