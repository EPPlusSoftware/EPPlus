/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Olap specific properties for the connection.
    /// </summary>
    public class ExcelConnectionOlapProperties
    {
        /// <summary>
        /// Indicates if we should get data from the local cube on refresh versus the original data source.
        /// </summary>
        public bool Local { get; set; } = false;
        /// <summary>
        /// The local connection string used when <see cref="Local"/> is true.
        /// </summary>
        public string LocalConnection { get; set; }
        /// <summary>
        /// If we should refresh the local cube from the original data source. When true, the original OLAP data source is queried each time the user explicitly refreshes the data in the application, and a new local cube is constructed from this query.
        /// </summary>
        public bool LocalRefresh { get; set; } = true;
        /// <summary>
        /// If true, the spreadsheet app should send the user interface locale ID to the OLAP provider to retrieve localized member names and properties, etc. When false, no locale ID is expected.
        /// </summary>
        public bool SendLocale { get; set; } = false;
        /// <summary>
        /// Maximum number of drill-through rows to return when the user drills through an aggregate value in a PivotTable.
        /// </summary>
        public int? RowDrillCount { get; set; }
        /// <summary>
        /// When true a PivotTable based on an OLAP source should format the data and aggregate cells in the PivotTable view using the background color from the OLAP source if this information is available.When false, OLAP server background fill colors are ignored, and standard formatting rules within the worksheet are followed.
        /// </summary>
        public bool ServerFill { get; set; } = true;
        /// <summary>
        /// When true a PivotTable based on an OLAP source should format the data and aggregate cells in the PivotTable view using the number format from the OLAP source.When false standard formatting rules within the worksheet are followed.
        /// </summary>
        public bool ServerNumberFormat { get; set; } = true;
        /// <summary>
        /// When true, a PivotTable based on OLAP source should format the data and aggregate cells in the PivotTable view using the font from the OLAP source. When false, formatting rules within the worksheet are followed.
        /// </summary>
        public bool ServerFont { get; set; } = true;
        /// <summary>
        /// When true a PivotTable based on OLAP source should format the data and aggregate cells in the PivotTable view using the font color from the OLAP source.When false standard formatting rules within the worksheet are followed.
        /// </summary>
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