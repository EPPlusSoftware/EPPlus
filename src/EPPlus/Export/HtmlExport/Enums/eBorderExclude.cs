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
    /// Exclude border properties in the css
    /// </summary>
    [Flags]
    public enum eBorderExclude
    {
        /// <summary>
        /// Exclude all border properties.
        /// </summary>
        All = 0x0F,
        /// <summary>
        /// Exclude top border properties
        /// </summary>
        Top = 0x01,
        /// <summary>
        /// Exclude bottom border properties
        /// </summary>
        Bottom = 0x02,
        /// <summary>
        /// Exclude left border properties
        /// </summary>
        Left = 0x04,
        /// <summary>
        /// Exclude right border properties
        /// </summary>
        Right = 0x08
    }
}
