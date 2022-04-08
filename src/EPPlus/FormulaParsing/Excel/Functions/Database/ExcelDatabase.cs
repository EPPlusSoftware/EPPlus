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
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    internal class ExcelDatabase
    {
        private readonly ExcelDataProvider _dataProvider;
        private readonly int _fromCol;
        private readonly int _toCol;
        private readonly int _fieldRow;
        private readonly int _endRow;
        private readonly string _worksheet;
        private int _rowIndex;
        private readonly List<ExcelDatabaseField> _fields = new List<ExcelDatabaseField>();

        public IEnumerable<ExcelDatabaseField> Fields
        {
            get { return _fields; }
        }

        public ExcelDatabase(ExcelDataProvider dataProvider, string range)
        {
            _dataProvider = dataProvider;
            var address = new ExcelAddressBase(range);
            _fromCol = address._fromCol;
            _toCol = address._toCol;
            _fieldRow = address._fromRow;
            _endRow = address._toRow;
            _worksheet = address.WorkSheetName;
            _rowIndex = _fieldRow;
            Initialize();
        }

        private void Initialize()
        {
            var fieldIx = 0;
            for (var colIndex = _fromCol; colIndex <= _toCol; colIndex++)
            {
                var nameObj = GetCellValue(_fieldRow, colIndex);
                var name = nameObj != null ? nameObj.ToString().ToLower(CultureInfo.InvariantCulture) : string.Empty;
                _fields.Add(new ExcelDatabaseField(name, fieldIx++));
            }
        }

        private object GetCellValue(int row, int col)
        {
            return _dataProvider.GetRangeValue(_worksheet, row, col);
        }

        public bool HasMoreRows
        {
            get { return _rowIndex < _endRow; }
        }

        public ExcelDatabaseRow Read()
        {
            var retVal = new ExcelDatabaseRow();
            _rowIndex++;
            foreach (var field in Fields)
            {
                var colIndex = _fromCol + field.ColIndex;
                var val = GetCellValue(_rowIndex, colIndex);
                retVal[field.FieldName] = val;
            }
            return retVal;
        }
    }
}
