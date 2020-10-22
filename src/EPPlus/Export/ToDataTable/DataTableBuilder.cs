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
        {
            Require.That(options).IsNotNull();
            Require.That(range).IsNotNull();
            _options = options;
            _range = range;
            _sheet = _range.Worksheet;
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;

        internal DataTable Build()
        {
            var dataTable = string.IsNullOrEmpty(_options.DataTableName) ? new DataTable() : new DataTable(_options.DataTableName);
            if(!string.IsNullOrEmpty(_options.DataTableNamespace))
            {
                dataTable.Namespace = _options.DataTableNamespace;
            }
            var columnOrder = 0;
            for (var col = _range.Start.Column; col <= _range.End.Column; col++)
            {
                var row = _range.Start.Row;
                var name = _options.ColumnNamePrefix + ++columnOrder;
                var columnIndex = columnOrder - 1;
                if(_options.FirstRowIsColumnNames)
                {
                    name = _sheet.Cells[row, col].Value?.ToString();
                    if (name == null) throw new InvalidOperationException(string.Format("First row contains an empty cell at index {0}", col - _range.Start.Column));
                    name = GetColumnName(name);
                }
                
                // find type
                while (_sheet.Cells[++row, col] == null && row <= _range.End.Row)
                    ;
                if (row == _range.End.Row && _sheet.Cells[row, col].Value == null) throw new InvalidOperationException(string.Format("Column with index {0} does not contain any values", col));
                var type = _sheet.Cells[row, col].Value.GetType();
                
                // check mappings
                if(_options.PredefinedMappingsOnly && !_options.Mappings.ContainsMapping( columnIndex))
                {
                    continue;
                }
                else if(_options.Mappings.ContainsMapping(columnIndex) && _options.Mappings.GetByRangeIndex(columnIndex).DataColumnType != null)
                {
                    type = _options.Mappings[columnIndex].DataColumnType;
                }
                dataTable.Columns.Add(name, type);
                if(!_options.Mappings.ContainsMapping(columnIndex))
                {
                    bool allowNull = !type.IsValueType || (Nullable.GetUnderlyingType(type) != null);
                    _options.Mappings.Add(columnOrder - 1, name, type, allowNull);
                }
                else if(_options.Mappings.GetByRangeIndex(columnIndex).DataColumnType == null)
                {
                    _options.Mappings.GetByRangeIndex(columnIndex).DataColumnType = type;
                }
            }
            return dataTable;
        }

        private string GetColumnName(string name)
        {
            switch(_options.NameParsingStrategy)
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
