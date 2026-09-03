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

using System;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Setting for rendering of picture drawings
    /// </summary>
    [Obsolete("Use HtmlDrawingSettings instead (Settings.Drawings).")]
    public class HtmlPictureSettings
    {
        HtmlDrawingSettings _drawingsSettings;
        internal HtmlPictureSettings(HtmlDrawingSettings drawingsSettings)
        {
            _drawingsSettings = drawingsSettings;
        }

        /// <summary>
        /// If picture drawings should be included in the html. Default is <see cref="ePictureInclude.Exclude"/>
        /// </summary>
        public ePictureInclude Include
        {
            get
            {
                return (ePictureInclude)_drawingsSettings.Include;
            }
            set
            {
                _drawingsSettings.Include = (eDrawingInclude)value;
            }
        }
        /// <summary>
        /// If the image should be added as absolut or relative in the css.
        /// </summary>
        public ePicturePosition Position 
        {
            get
            {
                return (ePicturePosition)_drawingsSettings.Position;
            }
            set
            {
                _drawingsSettings.Position = (eDrawingPosition)value;
            }
        }
        /// <summary>
        /// If the margin in pixels from the top corner should be used. 
        /// If this property is set to true, the cells vertical alignment will be set to 'top', 
        /// otherwise alignment will be set to middle.
        /// </summary>
        public bool AddMarginTop 
        {
            get
            {
                return _drawingsSettings.AddMarginTop;
            }
            set
            {
                _drawingsSettings.AddMarginTop = value;
            } 
        }
        /// <summary>
        /// If the margin in pixels from the left corner should be used.
        /// If this property is set to true, the cells text alignment will be set to 'left', 
        /// otherwise alignment will be set to center.
        /// </summary>
        public bool AddMarginLeft
        {
            get
            {
                return _drawingsSettings.AddMarginLeft;
            }
            set
            {
                _drawingsSettings.AddMarginLeft = value;
            }
        }        /// <summary>
                 /// If set to true the original size of the image is used, 
                 /// otherwise the size in the workbook is used. Default is false.
                 /// </summary>
        public bool KeepOriginalSize
        {
            get
            {
                return _drawingsSettings.KeepOriginalSizeOnPictures;
            }
            set
            {
                _drawingsSettings.KeepOriginalSizeOnPictures = value;
            }
        }
        /// <summary>
        /// Exclude settings 
        /// </summary>
        public PictureCssExclude CssExclude
        {
            get
            {
                return _drawingsSettings.PictureCssExclude;
            }
        }
        /// <summary>
        /// Adds the Blip name as Id for the img element in the HTML.
        /// Characters [A-Z][0-9]-_ are allowed. The first character allows [A-Z]_. 
        /// Other characters will be replaced with an hyphen (-).
        /// </summary>
        public bool AddNameAsId
        {
            get
            {
                return _drawingsSettings.AddNameAsId;
            }
            set
            {
                _drawingsSettings.AddNameAsId = value;
            }
        }
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
