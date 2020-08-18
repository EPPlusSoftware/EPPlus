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
    /// Worksheet hidden enumeration
    /// </summary>
    public enum eWorkSheetHidden
    {
        /// <summary>
        /// The worksheet is visible
        /// </summary>
        Visible,
        /// <summary>
        /// The worksheet is hidden but can be shown by the user via the user interface
        /// </summary>
        Hidden,
        /// <summary>
        /// The worksheet is hidden and cannot be shown by the user via the user interface
        /// </summary>
        VeryHidden
    }
}
