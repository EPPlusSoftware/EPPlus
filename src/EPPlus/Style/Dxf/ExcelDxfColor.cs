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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A color in a differential formatting record
    /// </summary>
    public class ExcelDxfColor : DxfStyleBase

    {
        internal ExcelDxfColor(ExcelStyles styles) : base(styles)
        {

        }
        /// <summary>
        /// Gets or sets a theme color
        /// </summary>
        public eThemeSchemeColor? Theme { get; set; }
        /// <summary>
        /// Gets or sets an indexed color
        /// </summary>
        public int? Index { get; set; }
        /// <summary>
        /// Gets or sets the color to automativ
        /// </summary>
        public bool? Auto { get; set; }
        /// <summary>
        /// Gets or sets the Tint value for the color
        /// </summary>
        public double? Tint { get; set; }
        /// <summary>
        /// Sets the color.
        /// </summary>
        public Color? Color { get; set; }
        /// <summary>
        /// The Id
        /// </summary>
        protected internal override string Id
        {
            get { return GetAsString(Theme) + "|" + GetAsString(Index) + "|" + GetAsString(Auto) + "|" + GetAsString(Tint) + "|" + GetAsString(Color==null ? "" : Color.Value.ToArgb().ToString("x")); }
        }
        /// <summary>
        /// Set the color of the drawing
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(Color color)
        {
            Theme = null;
            Auto = null;
            Index = null;
            Color = color;
        }
        /// <summary>
        /// Set the color of the drawing
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(eThemeSchemeColor color)
        {
            Color = null;
            Auto = null;
            Index = null;
            Theme = color;
        }
        /// <summary>
        /// Set the color of the drawing
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(ExcelIndexedColor color)
        {
            Color = null;
            Theme = null;
            Auto = null;
            Index = (int)color;
        }
        /// <summary>
        /// Set the color to automatic
        /// </summary>
        public void SetAuto()
        {
            Color = null;
            Theme = null;
            Index = null;
            Auto = true;
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfColor(_styles) { Theme = Theme, Index = Index, Color = Color, Auto = Auto, Tint = Tint };
        }
        /// <summary>
        /// If the object has a value
        /// </summary>
        protected internal override bool HasValue
        {
            get
            {
                return Theme != null ||
                       Index != null ||
                       Auto != null ||
                       Tint != null ||
                       Color != null;
            }
        }
        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            throw new NotImplementedException();
        }
    }
}
