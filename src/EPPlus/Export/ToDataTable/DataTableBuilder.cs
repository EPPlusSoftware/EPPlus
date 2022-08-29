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
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTableBuilder
    {
        public DataTableBuilder(ToDataTableOptions options, ExcelRangeBase range)
            : this(options, range, null) { }
        public DataTableBuilder(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
        {
            Require.That(options).IsNotNull();
            Require.That(range).IsNotNull();
            _options = options;
            _range = range;
            _sheet = _range.Worksheet;
            _dataTable = dataTable;
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;
        private DataTable _dataTable;

        internal DataTable Build()
        {
            var columnNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            if(_dataTable == null)
            {
                _dataTable = string.IsNullOrEmpty(_options.DataTableName) ? new DataTable() : new DataTable(_options.DataTableName);
            }
            if(!string.IsNullOrEmpty(_options.DataTableNamespace))
            {
                _dataTable.Namespace = _options.DataTableNamespace;
            }
            var columnOrder = 0;
            for (var col = _range.Start.Column; col <= _range.End.Column; col++)
            {
                var row = _range.Start.Row;
                var name = _options.ColumnNamePrefix + ++columnOrder;
                var origName = name;
                var columnIndex = columnOrder - 1;
                if(_options.Mappings.ContainsMapping(columnIndex))
                {
                    name = _options.Mappings.GetByRangeIndex(columnIndex).DataColumnName;
                }
                else if (_options.FirstRowIsColumnNames)
                {                    
                    name = _sheet.GetValue(row, col)?.ToString();
                    origName = name;
                    if (name == null) throw new InvalidOperationException(string.Format("First row contains an empty cell at index {0}", col - _range.Start.Column));
                    name = GetColumnName(name);
                }
                else
                {
                    row--;
                }
                if(columnNames.Contains(name))
                {
                    throw new InvalidOperationException($"Duplicate column name : {name}");
                }
                columnNames.Add(name);
                // find type
                while (_sheet.GetValue(++row, col) == null && row <= _range.End.Row)
                    ;
                var v = _sheet.GetValue(row, col);
                if (row == _range.End.Row && v == null) throw new InvalidOperationException(string.Format("Column with index {0} does not contain any values", col));
                var type = v == null ? typeof(Nullable) : v.GetType();

                // check mapping
                var mapping = _options.Mappings.GetByRangeIndex(columnIndex);
                if (_options.PredefinedMappingsOnly && mapping == null)
                {
                    continue;
                }
                else if (mapping != null)
                {
                    if(mapping.ColumnDataType != null)
                    {
                        type = mapping.ColumnDataType;
                    }
                    if(mapping.HasDataColumn && _dataTable.Columns[mapping.DataColumnName] == null)
                    {
                        _dataTable.Columns.Add(mapping.DataColumn);
                    }
                }

                if((mapping == null || !mapping.HasDataColumn) && _dataTable.Columns[name] == null)
                {
                    var column = _dataTable.Columns.Add(name, type);
                    column.Caption = origName;
                }

                if (!_options.Mappings.ContainsMapping(columnIndex))
                {
                    bool allowNull = !type.IsValueType || (Nullable.GetUnderlyingType(type) != null);
                    _options.Mappings.Add(columnOrder - 1, name, type, allowNull);
                }
                else if(_options.Mappings.GetByRangeIndex(columnIndex).ColumnDataType == null)
                {
                    _options.Mappings.GetByRangeIndex(columnIndex).ColumnDataType = type;
                }
            }
            HandlePrimaryKeys(_dataTable);
            return _dataTable;
        }

        private void HandlePrimaryKeys(DataTable dataTable)
        {
            var pk = new DataTablePrimaryKey(_options);
            if(pk.HasKeys)
            {
                var cols = new List<DataColumn>();
                foreach(var colObj in dataTable.Columns)
                {
                    var col = colObj as DataColumn;
                    if (col == null) continue;
                    if (pk.ContainsKey(col.ColumnName))
                    {
                        cols.Add(col);
                    }   
                }
                dataTable.PrimaryKey = cols.ToArray();
            }
        }

        private string GetColumnName(string name)
        {
            switch(_options.ColumnNameParsingStrategy)
            {
                case NameParsingStrategy.Preserve:
                    return name;
                case NameParsingStrategy.SpaceToUnderscore:
                    return name.Replace(" ", "_");
                case NameParsingStrategy.RemoveSpace:
                    return name.Replace(" ", string.Empty);
                default:
                    return name;
            }
        }
    }
}
