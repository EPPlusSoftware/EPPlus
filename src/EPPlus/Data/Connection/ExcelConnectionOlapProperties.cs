/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Data.Connection
{
    public class ExcelConnectionOlapProperties
    {
        public bool Local { get; set; } = false;
        public string LocalConnection { get; set; }
        public bool LocalRefresh { get; set; } = true;
        public bool SendLocale { get; set; } = false;
        public int? RowDrillCount { get; set; }
        public bool ServerFill { get; set; } = true;
        public bool ServerNumberFormat { get; set; } = true;
        public bool ServerFont { get; set; } = true;
        public bool ServerFontColor { get; set; } = true;

        /// <summary>
        /// A list of tables used in the data model for the connection. This property only applies when <see cref="ExcelConnection.Type"/> is set to <see cref="eConnectionDataSourceType.DataModelOLEDB"/>."/>
        /// </summary>
        public List<string> Tables { get; } = new List<string>();
        /// <summary>
        /// The OLE DB command text that is used by a Model Data Source OLE DB. This property only applies when <see cref="ExcelConnection.Type"/> is set to <see cref="eConnectionDataSourceType.DataModelOLEDB"/>."/>
        /// </summary>
        public string DataModelCommand { get; } 
    }
}