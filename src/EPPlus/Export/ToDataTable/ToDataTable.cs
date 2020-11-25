﻿/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class ToDataTable
    {
        public ToDataTable(ToDataTableOptions options, ExcelRangeBase range)
        {
            Require.That(options).IsNotNull();
            Require.That(range).IsNotNull();
            _options = options;
            _range = range;
        }

        private readonly ToDataTableOptions _options;
        private readonly ExcelRangeBase _range;

        public DataTable Execute()
        {
            var dataTable = new DataTableBuilder(_options, _range).Build();
            new DataTableExporter(_options, _range, dataTable).Export();
            return dataTable;
        }

        public DataTable Execute(DataTable dataTable)
        {
            Require.That(dataTable).IsNotNull();
            new DataTableMapper(_options, _range, dataTable).Map();
            new DataTableExporter(_options, _range, dataTable).Export();
            return dataTable;
        }
    }
}
