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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class ExcelLookupNavigator : LookupNavigator
    {
        private int _currentRow;
        private int _currentCol;
        private object _currentValue;
        IRangeInfo _range;
        private int _index;

        public ExcelLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
            : base(direction, arguments, parsingContext)
        {
            Initialize();
        }

        private void Initialize()
        {
            _index = 0;
            var factory = new RangeAddressFactory(ParsingContext.ExcelDataProvider);
            if (Arguments.RangeInfo == null)
            {
                _range = ParsingContext.ExcelDataProvider.GetRange(ParsingContext.Scopes.Current.Address.Worksheet, Arguments.RangeAddress);
            }
            else
            {
                _range = Arguments.RangeInfo;
            }
            _currentCol = _range.Address._fromCol;
            _currentRow = _range.Address._fromRow;
            SetCurrentValue();
        }

        private void SetCurrentValue()
        {
            _currentValue = _range.GetValue(_currentRow, _currentCol);
        }

        private bool HasNext()
        {
            if (Direction == LookupDirection.Vertical)
            {
                return _currentRow < _range.Address._toRow;
            }
            else
            {
                return _currentCol < _range.Address._toCol;
            }
        }

        public override int Index
        {
            get { return _index; }
        }

        public override bool MoveNext()
        {
            if (!HasNext()) return false;
            if (Direction == LookupDirection.Vertical)
            {
                _currentRow++;
            }
            else
            {
                _currentCol++;
            }
            _index++;
            SetCurrentValue();
            return true;
        }

        public override object CurrentValue
        {
            get { return _currentValue; }
        }

        public override object GetLookupValue()
        {
            var row = _currentRow;
            var col = _currentCol;
            if (Direction == LookupDirection.Vertical)
            {
                col += Arguments.LookupIndex - 1;
                row += Arguments.LookupOffset;
            }
            else
            {
                row += Arguments.LookupIndex - 1;
                col += Arguments.LookupOffset;
            }
            return _range.GetValue(row, col); 
        }
    }
}
