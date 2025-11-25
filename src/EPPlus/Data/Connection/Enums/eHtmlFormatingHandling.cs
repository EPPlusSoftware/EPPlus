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
    /// How to handle formatting from the HTML source when bringing web query data into the worksheet.
    /// </summary>
    public enum eHtmlFormatingHandling
    {
        /// <summary>
        /// No formatting is applied.
        /// </summary>
        None = 1,
        /// <summary>
        /// HTML formatting should be translated into rich text formatting when importing the data.
        /// </summary>
        RTF = 2,
        /// <summary>
        /// All HTML formatting should be preserved when importing the data.
        /// </summary>
        All = 3 
    }
}