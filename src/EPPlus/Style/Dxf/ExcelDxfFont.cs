/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/01/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A font in a differential formatting record
    /// </summary>
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
        /// <summary>
        /// Displays only the inner and outer borders of each character. Similar to bold
        /// </summary>
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
        /// <summary>
        /// Shadow for the font. Used on Macintosh only.
        /// </summary>
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
        /// <summary>
        /// Condence (squeeze it together). Used on Macintosh only.
        /// </summary>
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
        /// <summary>
        /// Extends or stretches the text. Legacy property used in older speadsheet applications.
        /// </summary>
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
        /// <summary>
        /// Which font scheme to use from the theme
        /// </summary>
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
        /// <summary>
        /// The Id to identify the font uniquely
        /// </summary>
        internal override string Id
        {
            get
            {
                return GetAsString(Bold) + "|" + GetAsString(Italic) + "|" + GetAsString(Strike) + "|" + (Color == null ? ExcelDxfColor.GetEmptyId() : Color.Id) + "|" + GetAsString(Underline)
                    + "|" + GetAsString(Name) + "|" + GetAsString(Size) + "|" + GetAsString(Family) + "|" + GetVAlign() + "|" + GetAsString(Outline) + "|" + GetAsString(Shadow) + "|" + GetAsString(Condense) + "|" + GetAsString(Extend) + "|" + GetAsString(Scheme);
            }
        }
        private string GetVAlign()
        {
            return VerticalAlign==ExcelVerticalAlignmentFont.None ? "" : GetAsString(VerticalAlign);
        }

        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        internal override DxfStyleBase Clone()
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
        /// <summary>
        /// If the object has any properties set
        /// </summary>
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
        /// <summary>
        /// Clears all properties
        /// </summary>
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
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            helper.CreateNode(path);
            SetValueBool(helper, path + "/d:b/@val", Bold);
            SetValueBool(helper, path + "/d:i/@val", Italic);
            SetValueBool(helper, path + "/d:strike/@val", Strike);
            SetValueBool(helper, path + "/d:condense/@val", Condense);
            SetValueBool(helper, path + "/d:extend/@val", Extend);
            SetValueBool(helper, path + "/d:outline/@val", Outline);
            SetValueBool(helper, path + "/d:shadow/@val", Shadow);
            SetValue(helper, path + "/d:u/@val", Underline == null ? null : Underline.ToEnumString());
            SetValue(helper, path + "/d:vertAlign/@val", VerticalAlign == ExcelVerticalAlignmentFont.None ? null : VerticalAlign.ToEnumString());
            SetValue(helper, path + "/d:sz/@val", Size);
            SetValueColor(helper, path + "/d:color", Color);
            SetValue(helper, path + "/d:name/@val", Name);
            SetValue(helper, path + "/d:family/@val", Family);
            SetValue(helper, path + "/d:scheme/@val", Scheme.HasValue ? Scheme.ToEnumString() : null);
        }
        internal override void SetValuesFromXml(XmlHelper helper)
        {
            Size = helper.GetXmlNodeIntNull("d:font/d:sz/@val");

            base.SetValuesFromXml(helper);

            Name = helper.GetXmlNodeString("d:font/d:name/@val");
            Condense = helper.GetXmlNodeBoolNullable("d:font/d:condense/@val");
            Extend = helper.GetXmlNodeBoolNullable("d:font/d:extend/@val");
            Outline = helper.GetXmlNodeBoolNullable("d:font/d:outline/@val");
            
            var v = helper.GetXmlNodeString("d:font/d:vertAlign/@val");
            VerticalAlign = string.IsNullOrEmpty(v)?ExcelVerticalAlignmentFont.None: v.ToEnum(ExcelVerticalAlignmentFont.None);
            
            Family = helper.GetXmlNodeIntNull("d:font/d:family/@val");
            Scheme = helper.GetXmlEnumNull<eThemeFontCollectionType>("d:font/d:scheme/@val");
            Shadow = helper.GetXmlNodeBoolNullable("d:font/d:shadow/@val");
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                base.SetStyle();
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Name, _name);
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Size, _size);
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Family, _family);
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.Scheme, _scheme);
                _callback?.Invoke(eStyleClass.Font, eStyleProperty.VerticalAlign, _verticalAlign);

            }
        }
    }
}