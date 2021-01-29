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
        internal ExcelDxfFontBase(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(styles, callback)
        {
            Color = new ExcelDxfColor(styles, eStyleClass.Font, callback);
        }
        bool? _bold;
        /// <summary>
        /// Font bold
        /// </summary>
        public bool? Bold
        {
            get
            {
                return _bold;
            }
            set
            {
                _bold = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Bold, value);
            }
        }
        bool? _italic;
        /// <summary>
        /// Font Italic
        /// </summary>
        public bool? Italic
        {
            get
            {
                return _italic;
            }
            set
            {
                _italic = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Italic, value);
            }
        }
        bool? _strike;
        /// <summary>
        /// Font-Strikeout
        /// </summary>
        public bool? Strike
        {
            get
            {
                return _strike;
            }
            set
            {
                _strike = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Strike, value);
            }
        }
        /// <summary>
        /// The color of the text
        /// </summary>
        public ExcelDxfColor Color { get; set; }
        ExcelUnderLineType? _underline;
        /// <summary>
        /// The underline type
        /// </summary>
        public ExcelUnderLineType? Underline
        {
            get
            {
                return _underline;
            }
            set
            {
                _underline = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.UnderlineType, value);
            }
        }

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
            return new ExcelDxfFontBase(_styles, _callback) { Bold = Bold, Color = (ExcelDxfColor)Color.Clone(), Italic = Italic, Strike = Strike, Underline = Underline };
        }

        internal void GetValuesFromXml(XmlHelperInstance helper)
        {
            if (helper.ExistsNode("d:font"))
            {
                Bold = helper.GetXmlNodeBoolNullableWithVal("d:font/d:b");
                Italic = helper.GetXmlNodeBoolNullableWithVal("d:font/d:i");
                Strike = helper.GetXmlNodeBoolNullableWithVal("d:font/d:strike");
                Underline = GetUnderLine(helper);
                Color = GetColor(helper, "d:font/d:color", eStyleClass.Font);
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
        internal ExcelDxfFont(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
           : base(styles, callback)
        {
        }
        float? _size;
        /// <summary>
        /// The font size 
        /// </summary>
        public float? Size
        {
            get
            {
                return _size;
            }
            set
            {
                _size = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Size, value);
            }
        }
        string _name;
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Name, value);
            }
        }
        int? _family;
        /// <summary>
        /// Font family 
        /// </summary>
        public int? Family
        {
            get
            {
                return _family;
            }
            set
            {
                _family = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Family, value);
            }
        }
        ExcelVerticalAlignmentFont _verticalAlign = ExcelVerticalAlignmentFont.None;
        /// <summary>
        /// Font-Vertical Align
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign
        {
            get
            {
                return _verticalAlign;
            }
            set
            {
                _verticalAlign = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.VerticalAlign, value);
            }
        }
        bool? _outline;
        public bool? Outline
        {
            get
            {
                return _outline;
            }
            set
            {
                _outline = value;                
            }
        }
        bool? _shadow;
        public bool? Shadow
        {
            get
            {
                return _shadow;
            }
            set
            {
                _shadow = value;
            }
        }
        bool? _condense;
        public bool? Condense 
        {
            get
            {
                return _condense;
            }
            set
            {
                _condense = value;
            }
        }
        bool? _extend;
        public bool? Extend
        {
            get
            {
                return _extend;
            }
            set
            {
                _extend = value;
            }
        }
        eThemeFontCollectionType? _scheme;
        public eThemeFontCollectionType? Scheme
        {
            get
            {
                return _scheme;
            }
            set
            {
                _scheme = value;
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Scheme, value);
            }
        }
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
            return new ExcelDxfFont(_styles, _callback) 
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