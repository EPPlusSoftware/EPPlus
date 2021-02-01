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
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                _callback.Invoke(eStyleClass.Font, eStyleProperty.Bold, _bold);
                _callback.Invoke(eStyleClass.Font, eStyleProperty.Italic, _italic);
                _callback.Invoke(eStyleClass.Font, eStyleProperty.Strike, _strike);
                _callback.Invoke(eStyleClass.Font, eStyleProperty.UnderlineType, Underline);
                Color.SetStyle();
            }
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
}