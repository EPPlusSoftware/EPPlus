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
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Base class for css export settings.
    /// </summary>
    public abstract class CssExportSettings
    {
        /// <summary>
        /// If set to true shared css classes used on table elements are included in the css. 
        /// If set to false, these classes has to be included manually. <see cref="IncludeNormalFont"/> will be ignored if set to false and no font css will be added.        
        /// Default is true
        /// </summary>
        public bool IncludeSharedClasses { get; set; } = true;
        /// <summary>
        /// If true the normal font will be included in the css. Default is true
        /// </summary>
        public bool IncludeNormalFont { get; set; } = true;

        /// <summary>
        /// Css elements added to the table.
        /// </summary>
        public Dictionary<string, string> AdditionalCssElements
        {
            get;
            internal set;
        }
        /// <summary>
        /// The value used in the stylesheet for an indentation in a cell
        /// </summary>
        public float IndentValue { get; set; } = 2;
        /// <summary>
        /// The unit used in the stylesheet for an indentation in a cell
        /// </summary>
        public string IndentUnit { get; set; } = "em";
        internal void ResetToDefaultInternal()
        {
            AdditionalCssElements = new Dictionary<string, string>()
            {
                { "border-spacing", "0" },
                { "border-collapse", "collapse" },
                { "word-wrap", "break-word"},
                { "white-space", "nowrap"},
            };
            IndentValue = 2;
            IndentUnit = "em";
        }
        internal void CopyInternal(CssExportSettings copy)
        {
            AdditionalCssElements = copy.AdditionalCssElements;
            IndentValue = copy.IndentValue;
            IndentUnit = copy.IndentUnit;
        }
    }
}
