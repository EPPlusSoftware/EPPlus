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
    /// Database specific properties.
    /// </summary>
    public class ExcelDatabaseProperties : IDocumentPartSave
    {
        /// <summary>
        /// A connection string used to initiate the connection.
        /// </summary>
        public string Connection { get; set; }
        /// <summary>
        /// The command to use. For example a table or a SQL statment.
        /// </summary>
        public string Command { get; set; }
        /// <summary>
        /// The type of command
        /// </summary>
        public eCommandType CommandType { get; set; }
        /// <summary>
        /// A second command text string that is persisted when PivotTable server-based page fields are in use.
        /// </summary>
        public string ServerCommand { get; set; }

        public void Save()
        {
            throw new System.NotImplementedException();
        }
    }
}