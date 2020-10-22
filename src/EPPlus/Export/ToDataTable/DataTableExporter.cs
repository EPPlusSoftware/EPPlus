/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Utilities;
using ConvertUtility = OfficeOpenXml.Utils.ConvertUtil;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTableExporter
    {
        public DataTableExporter(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
        {
            Require.That(options).IsNotNull();
            Require.That(range).IsNotNull();
            Require.That(dataTable).IsNotNull();
            _options = options;
            _range = range;
            _sheet = _range.Worksheet;
            _dataTable = dataTable;
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;
        private readonly DataTable _dataTable;

        public void Export()
        {
            var row = _options.FirstRowIsColumnNames ? _range.Start.Row + 1 : _range.Start.Row;
            while(row <= _range.End.Row)
            {
                var dataRow = _dataTable.NewRow();
                foreach (var mapping in _options.Mappings)
                {
                    var col = mapping.ZeroBasedColumnIndexInRange + _range.Start.Column;
                    var val = _sheet.GetValueInner(row, col);
                    if(mapping.DataColumnType == typeof(DateTime))
                    {
                        dataRow[mapping.DataColumnName] = ConvertUtility.GetValueDate(val);
                    }
                    else
                    {
                        dataRow[mapping.DataColumnName] = val;
                    }
                }
                _dataTable.Rows.Add(dataRow);
                row++;
            }
        }
    }
}
