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
using OfficeOpenXml.Core;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using System;
using System.IO;

namespace OfficeOpenXml.Data.Connection
{

    /// <summary>
    /// Which credentials to use.
    /// </summary>
    public enum eCredential
    {
        /// <summary>
        /// Use integrated security.
        /// </summary>
        Integrated,
        /// <summary>
        /// No credentials.
        /// </summary>
        None,
        /// <summary>
        /// Use stored credentials.
        /// </summary>
        Stored,
        /// <summary>
        /// Prompt for credentials.
        /// </summary>
        Prompt
    }
    /// <summary>
    /// How a connection object should reconnect. 
    /// </summary>
    public enum eReconnectionMethod
    {
        /// <summary>
        /// On refresh use the existing connection information. If the existing information cannot be used to establish a connection, get updated connection information, if available from the external connection file. 
        /// </summary>
        AsRequired = 1,
        /// <summary>
        /// On every refresh get updated connection information from the external connection file, if available, and use that instead of the existing connection information. In this case the data refresh will fail if the external connection file is unavailable.
        /// </summary>
        Always = 2,
        /// <summary>
        ///Never get updated connection information from the external connection file even if it is available and even if the existing connection information cannot be used. The possible values for this attribute are defined by the W3C XML Schema unsignedInt datatype. 
        /// </summary>
        Never = 3
    }
    /// <summary>
    /// Represents a connection to an external data source in a workbook.
    /// </summary>
    public class ExcelConnection : DocumentPart<ExcelConnection>
    {
        internal ExcelConnection(IDocumentPart<ExcelConnection> dp) : base(dp)
        {
        }
        /// <summary>
        /// Gets or sets the unique identifier for the connection.
        /// </summary>
        public int Id { get; set; } 
        /// <summary>
        /// Gets or sets the name associated with the connection. Each connection should have a unique name.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// A description of the connection.
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// How the connection should handle credentials.
        /// </summary>
        public eCredential Credentials { get; set; } = eCredential.Integrated;
        /// <summary>
        /// If the connection has been deleted.
        /// </summary>
        public bool IsDeleted { get; set; }
        /// <summary>
        /// If the connection can be refreshed in the background.
        /// </summary>
        public bool IsBackground { get; set; }
        /// <summary>
        /// Specifies the interval, in minutes, at which the connection is automatically refreshed. A value of 0 means that the connection is not automatically refreshed.
        /// </summary>
        public int AutomaticRefreshInterval { get; set; } = 0;
        /// <summary>
        /// If true, the connection should be kept alive by the spreadsheet application.
        /// </summary>
        public bool KeepAlive { get; set; } = false;
        /// <summary>
        /// The minimum version of the spreadsheet application that can refresh the connection.
        /// </summary>
        public int MinimumRefreshableVersion { get; set; } = 0;
        /// <summary>
        /// If the connection has been refreshed for the first time
        /// </summary>
        public bool IsNew { get; set; } = false;
        /// <summary>
        /// The full external path to the ODC file from which the connection was created.
        /// </summary>
        public string OdcFile { get; set; }
        /// <summary>
        /// If the spreadsheet application should always and only use the connection information in the external connection file indicated by the <see cref="OdcFile"/> when the connection is refreshed.
        /// </summary>
        public bool OnlyUseConnectionFile { get; set; } = false;
        /// <summary>
        /// The version of the spreadsheet application when the connection was last refreshed.
        /// </summary>
        public int LastRefreshVersion { get; set; }
        /// <summary>
        /// How the connection should reconnect when the connection fails
        /// </summary>
        public eReconnectionMethod ReconnectionMethod { get; set; } = eReconnectionMethod.AsRequired;
        /// <summary>
        /// If the connection should be refreshed when the workbook is loaded.
        /// </summary>
        public bool RefreshOnLoad { get; set; } = false;
        /// <summary>
        /// If data fetched by the connection should be saved in the workbook. Default false.
        /// </summary>
        public bool SaveData { get; set; } = false;
        /// <summary>
        /// If the password for the connection should be saved in the workbook. Default false.
        /// </summary>
        public bool SavePassword { get; set; } = false;
        /// <summary>
        /// SSO id used for authentication for the connection.
        /// </summary>
        public string SingleSignOnId { get; set; }
        /// <summary>
        /// Used when the external data source is file-based. When a connection to such a data source fails, the spreadsheet application attempts to connect directly to this file.Can be expressed in URI or system-specific file path notation.
        /// </summary>
        public string SourceDatabaseFile { get; set; }
        /// <summary>
        /// The type of data source for the connection.
        /// </summary>
        public eConnectionDataSourceType? Type { get; set; } = null;
        /// <summary>
        /// Database specific properties for the connection. This property is null if the connection is not a database or olap connection.
        /// </summary>
        public ExcelDatabaseProperties DatabaseProperties { get; internal set; }
        /// <summary>
        /// Olap specific properties for the connection. This property is null if the connection is not an olap connection.
        /// </summary>
        public ExcelConnectionOlapProperties OlapProperties { get; internal set; }
        /// <summary>
        /// Web specific properties for the connection. This property is null if the connection is not a web connection.
        /// </summary>
        public ExcelWebProperties WebProperties { get; internal set; }
        /// <summary>
        /// Text specific properties for the connection. This property is null if the conenction is not a text connection.
        /// </summary>
        public ExcelTextProperties TextProperties { get; internal set; }
        /// <summary>
        /// Parameters for the connection.
        /// </summary>
        public ExcelConnectionParameters Parameters { get; internal set; }

    }
}