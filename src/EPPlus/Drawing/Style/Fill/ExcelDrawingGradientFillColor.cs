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
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// Represents a color in the gradiant color list
    /// </summary>
    public class ExcelDrawingGradientFillColor
    {
        /// <summary>
        /// The position of color in a range from 0-100%
        /// </summary>
        public double Position { get; internal set; }
        /// <summary>
        /// The color to use.
        /// </summary>
        public ExcelDrawingColorManager Color { get; set; }
        internal XmlNode TopNode { get; set; }
    }
}
