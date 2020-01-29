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

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A base class for differential formatting font styles
    /// </summary>
    public class ExcelDxfFontBase : DxfStyleBase<ExcelDxfFontBase>
    {
        internal ExcelDxfFontBase(ExcelStyles styles)
            : base(styles)
        {
            Color = new ExcelDxfColor(styles);
        }
        /// <summary>
        /// Font bold
        /// </summary>
        public bool? Bold
        {
            get;
            set;
        }
        /// <summary>
        /// Font Italic
        /// </summary>
        public bool? Italic
        {
            get;
            set;
        }
        /// <summary>
        /// Font-Strikeout
        /// </summary>
        public bool? Strike { get; set; }
        //public float? Size { get; set; }
        /// <summary>
        /// The color of the text
        /// </summary>
        public ExcelDxfColor Color { get; set; }
        //public string Name { get; set; }
        //public int? Family { get; set; }
        ///// <summary>
        ///// Font-Vertical Align
        ///// </summary>
        //public ExcelVerticalAlignmentFont? VerticalAlign
        //{
        //    get;
        //    set;
        //}
        /// <summary>
        /// The underline type
        /// </summary>
        public ExcelUnderLineType? Underline { get; set; }

        /// <summary>
        /// The id
        /// </summary>
        protected internal override string Id
        {
            get
            {
                return GetAsString(Bold) + "|" + GetAsString(Italic) + "|" + GetAsString(Strike) + "|" + (Color ==null ? "" : Color.Id) + "|" /*+ GetAsString(VerticalAlign) + "|"*/ + GetAsString(Underline);
            }
        }

        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            helper.CreateNode(path);
            SetValueBool(helper, path + "/d:b/@val", Bold);
            SetValueBool(helper, path + "/d:i/@val", Italic);
            SetValueBool(helper, path + "/d:strike", Strike);
            SetValue(helper, path + "/d:u/@val", Underline);
            SetValueColor(helper, path + "/d:color", Color);
        }
        /// <summary>
        /// If the font has a value
        /// </summary>
        protected internal override bool HasValue
        {
            get
            {
                return Bold != null ||
                       Italic != null ||
                       Strike != null ||
                       Underline != null ||
                       Color.HasValue;
            }
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override ExcelDxfFontBase Clone()
        {
            return new ExcelDxfFontBase(_styles) { Bold = Bold, Color = Color.Clone(), Italic = Italic, Strike = Strike, Underline = Underline };
        }
    }
}
