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
    /// HTML Settings for excluding picture css settings.
    /// </summary>
    public class PictureCssExclude
    {
        internal PictureCssExclude()
        {

        }
        /// <summary>
        /// Exclude image border CSS
        /// </summary>
        public bool Border { get; set; }
        /// <summary>
        /// Exclude image alignment CSS
        /// </summary>
        public bool Alignment { get; set; }

        /// <summary>
        /// Reset the setting to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            Border = false;
            Alignment = false;
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(PictureCssExclude copy)
        {
            Border = copy.Border;
            Alignment = copy.Alignment;
        }
    }
}
