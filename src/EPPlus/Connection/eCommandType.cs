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

namespace OfficeOpenXml.Connection
{
    public enum eCommandType
    {
        /// <summary>
        /// The command specifies a cube name.
        /// </summary>
        Cube = 1,
        /// <summary>
        /// The command is a SQL statment.
        /// </summary>
        SqlStatement = 2, 
        /// <summary>
        /// The command is a table.
        /// </summary>
        Table = 3,
        /// <summary>
        /// The command is left to the provider to interpret.
        /// </summary>
        ProviderInterpreted = 4,
        /// <summary>
        ///  Query is against a web based List Data Provider.
        /// </summary>
        WebBasedListDataProvider = 5
    }
}