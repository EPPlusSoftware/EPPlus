/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// If the Picture is absolut or relative to the table cell
    /// </summary>
    public enum ePicturePosition
    {
        /// <summary>
        /// No CSS is added for Position
        /// </summary>
        DontSet,
        /// <summary>
        /// Position is Absolute in the CSS
        /// </summary>
        Absolute,
        /// <summary>
        /// Position is Relative in the CSS
        /// </summary>
        Relative
    }
}
