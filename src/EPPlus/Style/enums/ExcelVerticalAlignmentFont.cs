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
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Font-Vertical Align
    /// </summary>
    public enum ExcelVerticalAlignmentFont
    {
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// The text in the parent run will be located at the baseline and presented in the same size as surrounding text
        /// </summary>
        Baseline,
        /// <summary>
        /// The text will be subscript.
        /// </summary>
        Subscript,
        /// <summary>
        /// The text will be superscript.
        /// </summary>
        Superscript
    }
}