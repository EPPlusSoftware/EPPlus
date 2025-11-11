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
    /// OleDb properties for Data Model connections.
    /// </summary>
    public class ExcelDataModelOleDbProperties : ExcelDataModelDataFeedProperties
    {
        /// <summary>
        /// A command used for the connection. If this property is set, <see cref="ExcelDataModelDataFeedProperties.Tables"/> must be empty.
        /// </summary>
        public string Command { get; set; }
    }
}