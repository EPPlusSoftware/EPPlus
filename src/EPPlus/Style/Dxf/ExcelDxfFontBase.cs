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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A base class for differential formatting font styles
    /// </summary>
    public class ExcelDxfFontBase : DxfStyleBase
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
        /// <summary>
        /// The color of the text
        /// </summary>
        public ExcelDxfColor Color { get; set; }
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
                return GetAsString(Bold) + "|" + GetAsString(Italic) + "|" + GetAsString(Strike) + "|" + (Color == null ? "" : Color.Id) + "|" + GetAsString(Underline);
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
            SetValueBool(helper, path + "/d:strike/@val", Strike);
            SetValue(helper, path + "/d:u/@val", Underline==null?null:Underline.ToEnumString());
            SetValueColor(helper, path + "/d:color", Color);
        }
        /// <summary>
        /// If the font has a value
        /// </summary>
        public override bool HasValue
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
        public override void Clear()
        {
            Bold = null;
            Italic = null;
            Strike = null;
            Underline = null;
            Color.Clear();
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfFontBase(_styles) { Bold = Bold, Color = (ExcelDxfColor)Color.Clone(), Italic = Italic, Strike = Strike, Underline = Underline };
        }

        internal void GetValuesFromXml(XmlHelperInstance helper)
        {
            if (helper.ExistsNode("d:font"))
            {
                Bold = helper.GetXmlNodeBoolNullableWithVal("d:font/d:b");
                Italic = helper.GetXmlNodeBoolNullableWithVal("d:font/d:i");
                Strike = helper.GetXmlNodeBoolNullableWithVal("d:font/d:strike");
                Underline = GetUnderLine(helper);
                Color = GetColor(helper, "d:font/d:color");
            }
        }

        private ExcelUnderLineType? GetUnderLine(XmlHelperInstance helper)
        {
            if (helper.ExistsNode("d:font/d:u"))
            {
                var v = helper.GetXmlNodeString("d:font/d:u/@val");
                if (string.IsNullOrEmpty(v))
                {
                    return ExcelUnderLineType.Single;
                }
                else
                {
                    return GetUnderLineEnum(v);
                }
            }
            else
            {
                return null;
            }
        }
    }
    public class ExcelDxfFont : ExcelDxfFontBase
    {
        internal ExcelDxfFont(ExcelStyles styles)
           : base(styles)
        {
        }
        /// <summary>
        /// The font size 
        /// </summary>
        public float? Size { get; set; }
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Font family 
        /// </summary>
        public int? Family { get; set; }
        /// <summary>
        /// Font-Vertical Align
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign
        {
            get;
            set;
        } = ExcelVerticalAlignmentFont.None;
        public bool? Outline { get; set; }
        public bool? Shadow { get; set; }
        public bool? Condense { get; set; }
        public bool? Extend { get; set; }
        public eThemeFontCollectionType? Scheme { get; set; }
        protected internal override string Id
        {
            get
            {
                return base.Id + "|" + GetAsString(Name) + "|" + GetAsString(Size) + "|" + GetAsString(Family) + "|" + GetAsString(VerticalAlign) + "|" + GetAsString(Outline) + "|" + GetAsString(Shadow) + "|" + GetAsString(Condense)+ "|" + GetAsString(Extend)+ "|" + GetAsString(Scheme);
            }
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfFont(_styles) 
            {
                Name = Name,
                Size = Size,
                Family = Family,
                Bold = Bold, 
                Color = (ExcelDxfColor)Color.Clone(), 
                Italic = Italic, 
                Strike = Strike, 
                Underline = Underline,  
                Condense = Condense,
                Extend = Extend,
                Scheme=Scheme,
                Outline=Outline,
                Shadow=Shadow,
                VerticalAlign=VerticalAlign
            };
        }
        public override bool HasValue
        {
            get
            {
                return base.HasValue ||
                       string.IsNullOrEmpty(Name) == false ||
                       Size.HasValue ||
                       Family.HasValue ||
                       Condense.HasValue ||
                       Extend.HasValue ||
                       Scheme.HasValue ||
                       Outline.HasValue ||
                       Shadow.HasValue ||
                       VerticalAlign != ExcelVerticalAlignmentFont.None
;
            }
        }
        public override void Clear()
        {
            base.Clear();
            Name = null;
            Size = null;
            Family = null;
            Condense = null;
            Extend = null;
            Scheme = null;
            Outline = null;
            Shadow = null;
            VerticalAlign = ExcelVerticalAlignmentFont.None;
        }
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            base.CreateNodes(helper, path);
            SetValueBool(helper, path + "/d:condense/@val", Condense);
            SetValueBool(helper, path + "/d:extend/@val", Extend);
            SetValueBool(helper, path + "/d:outline/@val", Outline);
            SetValueBool(helper, path + "/d:shadow/@val", Shadow);
            SetValue(helper, path + "/d:name/@val", Name);
            SetValue(helper, path + "/d:size/@val", Size);
            SetValue(helper, path + "/d:family/@val", Family);
            SetValue(helper, path + "/d:vertAlign/@val", VerticalAlign==ExcelVerticalAlignmentFont.None ? null : VerticalAlign.ToEnumString());
        }
        internal new void GetValuesFromXml(XmlHelperInstance helper)
        {
            base.GetValuesFromXml(helper);
            Name = helper.GetXmlNodeString("d:font/d:name/@val");
            Size = helper.GetXmlNodeIntNull("d:font/d:sz/@val");
            Condense = helper.GetXmlNodeBoolNullable("d:font/d:condense/@val");
            Extend = helper.GetXmlNodeBoolNullable("d:font/d:extend/@val");
            Outline = helper.GetXmlNodeBoolNullable("d:font/d:outline/@val");
            
            var v = helper.GetXmlNodeString("d:font/d:vertAlign/@val");
            VerticalAlign = string.IsNullOrEmpty(v)?ExcelVerticalAlignmentFont.None: v.ToEnum(ExcelVerticalAlignmentFont.None);
            
            Family = helper.GetXmlNodeIntNull("d:font/d:family/@val");
            Scheme = helper.GetXmlEnumNull<eThemeFontCollectionType>("d:font/d:scheme/@val");
            Shadow = helper.GetXmlNodeBoolNullable("d:font/d:shadow/@val");
        }
    }
}