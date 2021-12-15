/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Drawing;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Information about the content, type, bounds and resolution of an image.
    /// </summary>
    public class ExcelImageInfo
    {
        internal ExcelImageInfo()
        {
        }
        /// <summary>
        /// The width of the image
        /// </summary>
        public double Width
        {
            get;
            internal set;
        }
        /// <summary>
        /// The height of the image
        /// </summary>
        public double Height
        {
            get;
            internal set;
        }
        /// <summary>
        /// The horizontal resolution of the image
        /// </summary>
        public double HorizontalResolution
        {
            get;
            internal set;
        } = ExcelDrawing.STANDARD_DPI;
        /// <summary>
        /// The vertical resolution of the image
        /// </summary>
        public double VerticalResolution
        {
            get;
            internal set;
        } = ExcelDrawing.STANDARD_DPI;
    }
}