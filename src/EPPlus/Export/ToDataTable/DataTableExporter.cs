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
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;
        private readonly DataTable _dataTable;
        private Dictionary<Type, MethodInfo> _convertMethods = new Dictionary<Type, MethodInfo>();

        public void Export()
        {
            var row = _options.FirstRowIsColumnNames ? _range.Start.Row + 1 : _range.Start.Row;
            Validate();
            row += _options.SkipNumberOfRowsStart;
            
            while (row <= (_range.End.Row - _options.SkipNumberOfRowsEnd))
            {
                var dataRow = _dataTable.NewRow();
                var ignoreRow = false;
                var rowIsEmpty = true;
                var rowErrorMsg = string.Empty;
                var rowErrorExists = false;
                foreach (var mapping in _options.Mappings)
                {
                    var col = mapping.ZeroBasedColumnIndexInRange + _range.Start.Column;
                    var val = _sheet.GetValueInner(row, col);
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
                    dataRow[mapping.DataColumnName] = CastToColumnDataType(val, type);
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
            }
        }

        private void Validate()
        {
            var startRow = _options.FirstRowIsColumnNames ? _range.Start.Row + 1 : _range.Start.Row;
            if (_options.SkipNumberOfRowsStart < 0 || _options.SkipNumberOfRowsStart > (_range.End.Row - startRow))
            {
                throw new IndexOutOfRangeException("SkipNumberOfRowsStart was out of range: " + _options.SkipNumberOfRowsStart);
            }
            if (_options.SkipNumberOfRowsEnd < 0 || _options.SkipNumberOfRowsEnd > (_range.End.Row - startRow))
            {
                throw new IndexOutOfRangeException("SkipNumberOfRowsEnd was out of range: " + _options.SkipNumberOfRowsEnd);
            }
            if((_options.SkipNumberOfRowsEnd + _options.SkipNumberOfRowsStart) > (_range.End.Row - startRow))
            {
                throw new ArgumentException("Total number of skipped rows was larger than number of rows in range");
            }
        }


        private object CastToColumnDataType(object val, Type dataColumnType)
        {
            if (val == null)
            {
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
                return ConvertUtility.GetValueDate(val);
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
