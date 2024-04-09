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
using System.Reflection;

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
            if(_options.AlwaysAllowNull)
            {
                foreach(var mapping in _options.Mappings)
                {
                    mapping.AllowNull = true;
                }
            }
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;
        private readonly DataTable _dataTable;
        private Dictionary<Type, MethodInfo> _convertMethods = new Dictionary<Type, MethodInfo>();

        public void Export()
        {
            var start1 = _range.Start.Column;
            var end1 = _range.End.Column;
            var start2 = _range.Start.Row;
            var end2 = _range.End.Row;
            if (_options.DataIsTransposed)
            {
                start2 = _range.Start.Column;
                end2 = _range.End.Column;
                start1 = _range.Start.Row;
                end1 = _range.End.Row;
            }

            var row = _options.FirstRowIsColumnNames ? start2 + 1 : start2;
            Validate();
            row += _options.SkipNumberOfRowsStart;
            
            while (row <= (end2 - _options.SkipNumberOfRowsEnd))
            {
                var dataRow = _dataTable.NewRow();
                dataRow.BeginEdit();
                var ignoreRow = false;
                var rowIsEmpty = true;
                var rowErrorMsg = string.Empty;
                var rowErrorExists = false;
                foreach (var mapping in _options.Mappings)
                {
                    var col = mapping.ZeroBasedColumnIndexInRange + start1;
                    var val = _options.DataIsTransposed ? _sheet.GetValue(col, row) : _sheet.GetValue(row, col);
                    if (val != null && rowIsEmpty) rowIsEmpty = false;
                    if(!mapping.AllowNull && val == null)
                    {
                        rowErrorMsg = $"Value cannot be null, row: {row}, col: {col}";
                        rowErrorExists = true;
                    }
                    else if(ExcelErrorValue.Values.IsErrorValue(val))
                    {
                        switch(_options.ExcelErrorParsingStrategy)
                        {
                            case ExcelErrorParsingStrategy.IgnoreRowWithErrors:
                                ignoreRow = true;
                                continue;
                            case ExcelErrorParsingStrategy.ThrowException:
                                throw new InvalidOperationException($"Excel error value {val.ToString()} detected at row: {row}, col: {col}");
                            default:
                                val = null;
                                break;
                        }
                    }
                    if(mapping.TransformCellValue != null)
                    {
                        val = mapping.TransformCellValue.Invoke(val);
                    }
                    var type = mapping.ColumnDataType;
                    if(type == null)
                    {
                        type = _dataTable.Columns[mapping.DataColumnName].DataType;
                    }
                    dataRow[mapping.DataColumnName] = CastToColumnDataType(val, type, mapping.AllowNull);
                }
                if(rowIsEmpty)
                {
                    if(_options.EmptyRowStrategy == EmptyRowsStrategy.StopAtFirst)
                    {
                        row++;
                        break;
                    }
                }
                else
                {
                    if(rowErrorExists)
                    {
                        throw new InvalidOperationException(rowErrorMsg);
                    }
                    if (!ignoreRow)
                    {
                        _dataTable.Rows.Add(dataRow);
                    }
                }
                row++;
                dataRow.EndEdit();
            }
        }

        private void Validate()
        {
            var start2 = _range.Start.Row;
            var end2 = _range.End.Row;
            if (_options.DataIsTransposed)
            {
                start2 = _range.Start.Column;
                end2 = _range.End.Column;
            }

            var startRow = _options.FirstRowIsColumnNames ? start2 + 1 : start2;
            if (_options.SkipNumberOfRowsStart < 0 || _options.SkipNumberOfRowsStart > (end2 - startRow))
            {
                throw new IndexOutOfRangeException("SkipNumberOfRowsStart was out of range: " + _options.SkipNumberOfRowsStart);
            }
            if (_options.SkipNumberOfRowsEnd < 0 || _options.SkipNumberOfRowsEnd > (end2 - startRow))
            {
                throw new IndexOutOfRangeException("SkipNumberOfRowsEnd was out of range: " + _options.SkipNumberOfRowsEnd);
            }
            if((_options.SkipNumberOfRowsEnd + _options.SkipNumberOfRowsStart) > (end2 - startRow))
            {
                throw new ArgumentException("Total number of skipped rows was larger than number of rows in range");
            }
        }


        private object CastToColumnDataType(object val, Type dataColumnType, bool allowNull)
        {
            if (val == null)
            {
                if (allowNull) return DBNull.Value;
                if (dataColumnType.IsValueType)
                {
                    return Activator.CreateInstance(dataColumnType);
                }
                return null;
            }
            if (val.GetType() == dataColumnType)
            {
                return val;
            }
            else if (dataColumnType == typeof(DateTime))
            {
                var date = ConvertUtility.GetValueDate(val);
				if(!date.HasValue) return DBNull.Value;
                return date.Value;
            }
            else if (dataColumnType == typeof(double))
            {
                return ConvertUtility.GetValueDouble(val);
            }
            else
            {
                try
                {
                    if(!_convertMethods.ContainsKey(dataColumnType))
                    {
                        MethodInfo methodInfo = typeof(ConvertUtility).GetMethod(nameof(ConvertUtility.GetTypedCellValue));
                        _convertMethods.Add(dataColumnType, methodInfo.MakeGenericMethod(dataColumnType));
                    }
                    var getTypedCellValue = _convertMethods[dataColumnType];
                    return getTypedCellValue.Invoke(null, new object[] { val });
                }
                catch
                {
                    if (dataColumnType.IsValueType)
                    {
                        return Activator.CreateInstance(dataColumnType);
                    }
                    return null;
                }
            }
        }
    }
}
