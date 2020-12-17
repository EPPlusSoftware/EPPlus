/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromDataTable
    {
        public LoadFromDataTable(ExcelRangeBase range, DataTable dataTable, LoadFromDataTableParams parameters)
        {
            _range = range;
            _worksheet = range.Worksheet;
            _dataTable = dataTable;
            _printHeaders = parameters.PrintHeaders;
            _tableStyle = parameters.TableStyle;
        }

        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _worksheet;
        private readonly DataTable _dataTable;
        private readonly bool _printHeaders;
        private readonly TableStyles _tableStyle;

        public ExcelRangeBase Load()
        {
            if (_dataTable == null)
            {
                throw (new ArgumentNullException("Table can't be null"));
            }

            if (_dataTable.Rows.Count == 0 && _printHeaders == false)
            {
                return null;
            }

            //var rowArray = new List<object[]>();
            var row = _range._fromRow;
            if (_printHeaders)
            {
                _worksheet._values.SetValueRow_Value(_range._fromRow, _range._fromCol, _dataTable.Columns.Cast<DataColumn>().Select((dc) => { return dc.Caption; }).ToArray());
                row++;
            }
            foreach (DataRow dr in _dataTable.Rows)
            {
                _range.Worksheet._values.SetValueRow_Value(row++, _range._fromCol, dr.ItemArray);
            }
            if (row != _range._fromRow) row--;

            // set table style
            int rows = (_dataTable.Rows.Count == 0 ? 1 : _dataTable.Rows.Count) + (_printHeaders ? 1 : 0);
            if (rows >= 0 && _dataTable.Columns.Count > 0 && _tableStyle != TableStyles.None)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_range._fromRow, _range._fromCol, _range._fromRow + rows - 1, _range._fromCol + _dataTable.Columns.Count - 1), _dataTable.TableName);
                tbl.ShowHeader = _printHeaders;
                tbl.TableStyle = _tableStyle;
            }

            return _worksheet.Cells[_range._fromRow, _range._fromCol, row, _range._fromCol + _dataTable.Columns.Count - 1];
        }
    }
}
