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
using System;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exclude font properties in the css
    /// </summary>
    [Flags]
    public enum eFontExclude
    {
        /// <summary>
        /// Exclude all font properties.
        /// </summary>
        All = 0x4F,
        /// <summary>
        /// Exclude the font name property
        /// </summary>
        Name = 0x01,
        /// <summary>
        /// Exclude the font size property
        /// </summary>
        Size = 0x02,
        /// <summary>
        /// Exclude the font color property
        /// </summary>
        Color = 0x04,
        /// <summary>
        /// Exclude the font bold property
        /// </summary>
        Bold = 0x08,
        /// <summary>
        /// Exclude the font italic property
        /// </summary>
        Italic = 0x10,
        /// <summary>
        /// Exclude the font strike property
        /// </summary>
        Strike = 0x20,
        /// <summary>
        /// Exclude the font underline property
        /// </summary>
        Underline = 0x40,
    }
}
