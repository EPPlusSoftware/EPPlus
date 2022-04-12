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
    /// How hidden rows are handled.
    /// </summary>
    public enum eHiddenState
    {
        /// <summary>
        /// Exclude hidden rows
        /// </summary>
        Exclude,
        /// <summary>
        /// Include hidden rows, but hide them.
        /// </summary>
        IncludeButHide,
        /// <summary>
        /// Include hidden rows.
        /// </summary>
        Include
    }
}
