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
using System;

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
        public int Id { get; set; } 
        public string Name { get; set; }
        public string Description { get; set; }
        public eCredential Credentials { get; set; } = eCredential.Integrated;
        public bool IsDeleted { get; set; } 
        public bool IsBackground { get; set; }
        public int AutomaticRefreshInterval { get; set; } = 0;
        public bool KeepAlive { get; set; } = false;
        public int MinimumRefreshableVersion { get; set; } = 0;
        public bool IsNew { get; set; } = false;
        public string OdcFile { get; set; }
        public bool OnlyUseConnectionFile { get; set; } = false;
        public int LastRefreshVersion { get; set; }
        public eReconnectionMethod ReconnectionMethod { get; set; } = eReconnectionMethod.AsRequired;
        public bool RefreshOnLoad { get; set; } = false;
        public bool SaveData { get; set; } = false;
        public bool SavePassword { get; set; } = false;
        public string SingleSignOnId { get; set; }
        public string SourceDatabaseFile { get; set; }
        public eConnectionDataSourceType? Type { get; set; } = null;
        public ExcelDatabaseProperties DatabaseProperties { get; internal set; }
        public ExcelConnectionOlapProperties OlapProperties { get; internal set; }
        public ExcelWebProperties WebProperties { get; internal set; }
        public ExcelTextProperties TextProperties { get; internal set; }
        public ExcelConnectionParameters Parameters { get; internal set; }

    }
}