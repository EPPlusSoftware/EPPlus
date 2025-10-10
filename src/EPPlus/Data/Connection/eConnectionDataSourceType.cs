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
    /// The type of data source for a connection.
    /// </summary>
    public enum eConnectionDataSourceType
    {
        /// <summary>
        /// An ODBC data source.
        /// </summary>
        ODBC = 1,
        /// <summary>
        /// A DAO data source.
        /// </summary>
        DAO = 2,
        /// <summary>
        /// An application-defined data source.
        /// </summary>
        ApplicationDefined = 3,
        /// <summary>
        /// A web query data source.
        /// </summary>
        WebQuery = 4,
        /// <summary>
        /// An OLEDB data source.
        /// </summary>
        OLEDB = 5,
        /// <summary>
        /// A text file data source.
        /// </summary>
        Text = 6,
        /// <summary>
        /// An ADO data source.
        /// </summary>
        ADO = 7,
        /// <summary>
        /// A data source for a Microsoft Data Service Provider (MSDSP).
        /// </summary>
        DSP = 8
    }
}