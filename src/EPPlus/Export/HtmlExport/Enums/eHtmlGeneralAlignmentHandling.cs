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
    /// How the text alignment is handled when the style is set to General
    /// </summary>
    public enum eHtmlGeneralAlignmentHandling
    {
        /// <summary>
        /// Dont set any alignment when alignment is set to general
        /// </summary>
        DontSet,
        /// <summary>
        /// If the column data type is numeric or date, alignment will be right otherwise left.
        /// </summary>
        ColumnDataType,
        /// <summary>
        /// If the cell value data type is numeric or date, alignment will be right otherwise left.
        /// </summary>
        CellDataType
    }
}
