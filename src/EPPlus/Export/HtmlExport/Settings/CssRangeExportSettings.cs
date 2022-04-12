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
    /// Settings for css export for tables
    /// </summary>
    public class CssRangeExportSettings : CssExportSettings
    {
        internal CssRangeExportSettings()
        {
            ResetToDefault();
        }
        /// <summary>
        /// Settings to exclude specific styles from the css.
        /// </summary>
        public CssExclude CssExclude { get; } = new CssExclude();
        /// <summary>
        /// Reset the settings to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            CssExclude.ResetToDefault();
            base.ResetToDefaultInternal();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(CssRangeExportSettings copy)
        {
            CssExclude.Copy(copy.CssExclude);
            base.CopyInternal(copy);
        }
    }
}
