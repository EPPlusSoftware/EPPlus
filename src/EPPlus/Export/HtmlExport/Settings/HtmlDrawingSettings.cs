using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Drawing handler
    /// </summary>
    public class HtmlDrawingSettings
    {
        internal HtmlDrawingSettings()
        {

        }
        /// <summary>
        /// Optional handle to set individual settings for a drawing. Returning null will use the default settings.
        /// </summary>
        public Func<ExcelDrawing, HtmlDrawingSettings> IndividualDrawingHandler { get; set; } = null;
        /// <summary>
        /// Option to handle if a drawing should be excluded or not.
        /// </summary>
        public Func<ExcelDrawing, bool> ExcludeDrawingHandler { get; set; } = null;
        /// <summary>
        /// If a drawing should be included in the export or not.
        /// </summary>
        public eDrawingInclude Include = eDrawingInclude.Exclude;
        /// <summary>
        /// If the drawing image should be added as absolut or relative in the css.
        /// </summary>
        public eDrawingPosition Position { get; set; } = eDrawingPosition.Relative;
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
        public bool KeepOriginalSizeOnPictures { get; set; } = false;
        /// <summary>
        /// Exclude settings 
        /// </summary>
        public PictureCssExclude PictureCssExclude { get; } = new PictureCssExclude();
        /// <summary>
        /// Adds the Blip name as Id for the img element in the HTML.
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
            Include = eDrawingInclude.Exclude;
            Position = eDrawingPosition.Relative;
            AddMarginLeft = false;
            AddMarginTop = false;
            KeepOriginalSizeOnPictures = false;
            PictureCssExclude.ResetToDefault();
        }
        /// <summary>
        /// Copy the values from another settings object.
        /// </summary>
        /// <param name="copy">The object to copy.</param>
        public void Copy(HtmlDrawingSettings copy)
        {
            Include = copy.Include;
            Position = copy.Position;
            AddMarginLeft = copy.AddMarginLeft;
            AddMarginTop = copy.AddMarginTop;
            KeepOriginalSizeOnPictures = copy.KeepOriginalSizeOnPictures;
            PictureCssExclude.Copy(copy.PictureCssExclude);
        }
    }
}
