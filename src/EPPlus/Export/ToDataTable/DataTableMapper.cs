using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.RichData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using System.Xml.Linq;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTableMapper
    {
        public DataTableMapper(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
        {
            Require.That(options).IsNotNull();
            Require.That(dataTable).IsNotNull();
            Require.That(range).IsNotNull();
            _options = options;
            _dataTable = dataTable;
            _range = range;
        }

        private readonly ToDataTableOptions _options;
        private readonly DataTable _dataTable;
        private readonly ExcelRangeBase _range;

        internal void Map()
        {
            var indexInRange = 0;
            foreach(var columnObj in _dataTable.Columns)
            {
                var column = columnObj as DataColumn;
                if (column == null) continue;
                if(!_options.Mappings.Any(x => string.Compare(column.ColumnName, x.DataColumnName, StringComparison.OrdinalIgnoreCase) == 0))
                {
                    if(_options.FirstRowIsColumnNames)
                    {
                        var ix = FindIndexInRange(column.ColumnName);
                        if (ix == -1) throw new InvalidOperationException("Column name not found in range: " + column.ColumnName);
                        _options.Mappings.Add(ix, column.ColumnName, column.DataType, column.AllowDBNull);
                    }
                    else
                    {
                        _options.Mappings.Add(indexInRange, column.ColumnName, column.DataType, column.AllowDBNull);
                    }
                    indexInRange++;
                }
            }
        }

        private int FindIndexInRange(string columnName)
        {
            var row = _range.Start.Row;
            var index = 0;
            for(var col = _range.Start.Column; col <= _range.End.Column; col++)
            {
                var cellVal = _range.Worksheet.GetValue(row, col);

                if (cellVal == null) continue;
                if (string.Compare(columnName, cellVal.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return index;
                }
                index++;
            }
            return -1;
        }
    }
}
