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

        public void Export()
        {
            var row = _options.FirstRowIsColumnNames ? _range.Start.Row + 1 : _range.Start.Row;
            while (row <= _range.End.Row)
            {
                var dataRow = _dataTable.NewRow();
                foreach (var mapping in _options.Mappings)
                {
                    var col = mapping.ZeroBasedColumnIndexInRange + _range.Start.Column;
                    var val = _sheet.GetValueInner(row, col);
                    dataRow[mapping.DataColumnName] = CastToColumnDataType(val, mapping.DataColumnType);
                }
                _dataTable.Rows.Add(dataRow);
                row++;
            }
        }

        private object CastToColumnDataType(object val, Type dataColumnType)
        {
            if (val.GetType() == dataColumnType)
            {
                return val;
            }
            else if (val == null)
            {
                if (dataColumnType.IsValueType)
                {
                    return Activator.CreateInstance(dataColumnType);
                }
                return null;
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
                    MethodInfo methodInfo = typeof(ConvertUtility).GetMethod(nameof(ConvertUtility.GetTypedCellValue));
                    MethodInfo genericMethod = methodInfo.MakeGenericMethod(dataColumnType);
                    return genericMethod.Invoke(null, new object[] { val });
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
