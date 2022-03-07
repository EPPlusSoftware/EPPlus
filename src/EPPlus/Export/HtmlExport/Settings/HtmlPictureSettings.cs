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
    /// Setting for rendering of picture drawings
    /// </summary>
    public class HtmlPictureSettings
    {
        internal HtmlPictureSettings()
        {

        }
        /// <summary>
        /// If picture drawings should be included in the html. Default is <see cref="ePictureInclude.Exclude"/>
        /// </summary>
        public ePictureInclude Include { get; set; } = ePictureInclude.Exclude;
        /// <summary>
        /// If the image should be added as absolut or relative in the css.
        /// </summary>
        public ePicturePosition Position { get; set; } = ePicturePosition.Relative;
        /// <summary>
        /// If the margin in pixels from the top corner should be used. 
        /// If this property is set to true, the cells vertical alignment will be set to 'top', 
        /// otherwise alignment will be set to middle.
        /// </summary>
        public bool AddMarginTop { get; set; } = false;
        /// <summary>
        /// If the margin in pixels from the left corner should be used.
        /// If this property is set to true, the cells text alignment will be set to 'left', 
        /// otherwise alignment will be set to center.
        /// </summary>
        public bool AddMarginLeft { get; set; } = false;
        /// <summary>
        /// If set to true the original size of the image is used, 
        /// otherwise the size in the workbook is used. Default is false.
        /// </summary>
        public bool KeepOriginalSize { get; set; } = false;
        /// <summary>
        /// Exclude settings 
        /// </summary>
        public PictureCssExclude CssExclude { get; } = new PictureCssExclude();
        /// <summary>
        /// Adds the Picture name as Id for the img element in the HTML.
        /// Characters [A-Z][0-9]-_ are allowed. The first character allows [A-Z]_. 
        /// Other characters will be replaced with an hyphen (-).
        /// </summary>
        public bool AddNameAsId
        {
            get;
            set;
        } = true;
        /// <summary>
        /// Reset the setting to it's default values.
        /// </summary>
        public void ResetToDefault()
        {
            Include = ePictureInclude.Exclude;
            Position = ePicturePosition.Relative;
            AddMarginLeft = false;
            AddMarginTop = false;
            KeepOriginalSize = false;
            CssExclude.ResetToDefault();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(HtmlPictureSettings copy)
        {
            Include = copy.Include;
            Position = copy.Position;
            AddMarginLeft = copy.AddMarginLeft;
            AddMarginTop = copy.AddMarginTop;
            KeepOriginalSize = copy.KeepOriginalSize;
            CssExclude.Copy(copy.CssExclude);
        }
    }
}
