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
using OfficeOpenXml.Table;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exclude css on an <see cref="ExcelTable"/>.
    /// </summary>
    public class CssExcludeStyle
    {
        internal CssExcludeStyle()
        {

        }
        /// <summary>
        /// Css settings for table styles
        /// </summary>
        public CssExclude TableStyle { get; } = new CssExclude();
        /// <summary>
        /// Css settings for cell styles.
        /// </summary>
        public CssExclude CellStyle { get; } = new CssExclude();
    }
}
