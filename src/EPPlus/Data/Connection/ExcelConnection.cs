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
namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Represents a connection to an external data source in a workbook.
    /// </summary>
    public class ExcelConnection
    {
        public int Id { get; set; } 
        public string Name { get; set; }
        public string Description { get; set; }
        public string ConnectionString { get; set; }
        public string CommandText { get; set; }
        public string CommandType { get; set; }
        public bool IsDeleted { get; set; } 
        public bool IsBackground { get; set; }
        public int AutomaticRefreshInterval { get; set; }
        public bool KeepAlive { get; set; }
        public int MinimumRefreshableVersion { get; set; }
        public bool IsNew { get; set; }
        public string OdcFile { get; set; }
        public bool OnlyUseConnectionFile { get; set; }
        public bool LastRefreshVersion { get; set; }
        public bool RefreshOnLoad { get; set; }
        public bool SaveData { get; set; }
        public bool SavePassword { get; set; }
        public string SingleSignOnId { get; set; }
        public string SourceDatabaseFile { get; set; }
        public eConnectionDataSourceType Type { get; }
        public ExcelDatabaseProperties DatabaseProperties { get; }
        public ExcelConnectionOlapProperties OlapProperties { get; }
        public ExcelWebProperties WebProperties { get; }
        public ExcelTextProperties TextProperties { get; }
        public ExcelConnectionParameters Parameters { get; }

    }
}