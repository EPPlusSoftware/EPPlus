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
using System.Collections.Generic;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Data model data feed specific properties.
    /// </summary>
    public class ExcelDataModelDataFeedProperties
    {
        /// <summary>
        /// The connection string for the data feed.
        /// </summary>
        public string Connection { get; set; }
        /// <summary>
        /// A list of tables in the data feed.
        /// </summary>
        public List<string> Tables { get; } = new List<string>();
    }
}