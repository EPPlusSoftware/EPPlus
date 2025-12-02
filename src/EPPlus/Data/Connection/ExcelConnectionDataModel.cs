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
using System;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Data model specific properties for the connection.
    /// </summary>
    public class ExcelConnectionDataModel
    {
        /// <summary>
        /// The identifier of the Data Model data source.
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        ///  If this connection is a connection to the spreadsheet data model. If "true", the <see cref="ExcelConnection.Type"/> property of the ancestor connection, MUST be equal to OLEDB.
        /// </summary>
        public bool IsModel { get; set; }
        /// <summary>
        /// If headers are included when importing text data in a data model. This property only applies when <see cref="Type"/> is set to <see cref="eConnectionDataSourceType.DataModelText"/>.
        /// </summary>
        public bool ModelTextHeaders { get; internal set; }
        /// <summary>
        /// If <see cref="Type"/> is set to <see cref="eConnectionDataSourceType.DataModelWorksheetData"/>, this is the source name, otherwise this property is ignored.
        /// </summary>
        public string RangeSourceName { get; set; }
        /// <summary>
        /// Data model data feed properties for the connection. This property is null if the connection is not a data feed data model connection.
        /// </summary>
        public ExcelDataModelDataFeedProperties DataFeedProperties { get; internal set; }

        /// <summary>
        /// OleDb properties used for a data model. This property is null if the connection is not a data model OleDb connection.
        /// </summary>
        public ExcelDataModelOleDbProperties OleDbProperties { get; internal set; }

    }
}