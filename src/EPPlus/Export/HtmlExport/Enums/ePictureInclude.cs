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
    /// How to include picture drawings in the html
    /// </summary>
    public enum ePictureInclude
    {
        /// <summary>
        /// Do not include pictures in the html export. Default
        /// </summary>
        Exclude,
        /// <summary>
        /// Include in css only, so they images can be added manually. 
        /// </summary>
        IncludeInCssOnly,
        /// <summary>
        /// Include the images in the html export.
        /// </summary>
        Include
    }
}
