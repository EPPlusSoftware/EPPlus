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
}