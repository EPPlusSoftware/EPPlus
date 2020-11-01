using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTableMapper
    {
        public DataTableMapper(ToDataTableOptions options, DataTable dataTable)
        {
            Require.That(options).IsNotNull();
            Require.That(dataTable).IsNotNull();
            _options = options;
            _dataTable = dataTable;
        }

        private readonly ToDataTableOptions _options;
        private readonly DataTable _dataTable;

        internal void Map()
        {
            var indexInRange = 0;
            foreach(var columnObj in _dataTable.Columns)
            {
                var column = columnObj as DataColumn;
                if (column == null) continue;
                if(!_options.Mappings.Any(x => string.Compare(column.ColumnName, x.DataColumnName, StringComparison.OrdinalIgnoreCase) == 0))
                {
                    _options.Mappings.Add(indexInRange, column.ColumnName, column.DataType, column.AllowDBNull);
                    indexInRange++;
                }
            }
        }
    }
}
