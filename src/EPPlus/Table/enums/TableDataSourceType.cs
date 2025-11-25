/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// Specifies what the table data is based on.
    /// </summary>
    public enum TableDataSourceType
    {
        /// <summary>
        /// The table is based on a worksheet data range.
        /// </summary>
        Worksheet = 0,
        /// <summary>
        /// The table is based on an XML mapping.
        /// </summary>
        Xml = 1,
        /// <summary>
        /// The table is based on an external data query.
        /// </summary>
        QueryTable = 2
    }
}