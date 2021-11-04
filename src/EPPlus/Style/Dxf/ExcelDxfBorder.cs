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
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// The border style of a drawing in a differential formatting record
    /// </summary>
    public class ExcelDxfBorderBase : DxfStyleBase
    {
        internal ExcelDxfBorderBase(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(styles, callback)
        {
            Left = new ExcelDxfBorderItem(_styles, eStyleClass.BorderLeft, callback);
            Right = new ExcelDxfBorderItem(_styles, eStyleClass.BorderRight, callback);
            Top = new ExcelDxfBorderItem(_styles, eStyleClass.BorderTop, callback);
            Bottom = new ExcelDxfBorderItem(_styles, eStyleClass.BorderBottom, callback);
            Vertical = new ExcelDxfBorderItem(_styles, eStyleClass.Border, callback);
            Horizontal = new ExcelDxfBorderItem(_styles, eStyleClass.Border, callback);
        }
        /// <summary>
        /// Left border style
        /// </summary>
        public ExcelDxfBorderItem Left
        {
            get;
            internal set;
        }
        /// <summary>
        /// Right border style
        /// </summary>
        public ExcelDxfBorderItem Right
        {
            get;
            internal set;
        }
        /// <summary>
        /// Top border style
        /// </summary>
        public ExcelDxfBorderItem Top
        {
            get;
            internal set;
        }
        /// <summary>
        /// Bottom border style
        /// </summary>
        public ExcelDxfBorderItem Bottom
        {
            get;
            internal set;
        }
        /// <summary>
        /// Horizontal border style
        /// </summary>
        public ExcelDxfBorderItem Horizontal
        {
            get;
            internal set;
        }
        /// <summary>
        /// Vertical border style
        /// </summary>
        public ExcelDxfBorderItem Vertical
        {
            get;
            internal set;
        }

        /// <summary>
        /// The Id
        /// </summary>
        protected internal override string Id
        {
            get
            {
                return Top.Id + Bottom.Id + Left.Id + Right.Id + Vertical.Id + Horizontal.Id;
            }
        }

        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            Left.CreateNodes(helper, path + "/d:left");
            Right.CreateNodes(helper, path + "/d:right");
            Top.CreateNodes(helper, path + "/d:top");
            Bottom.CreateNodes(helper, path + "/d:bottom");
            Vertical.CreateNodes(helper, path + "/d:vertical");
            Horizontal.CreateNodes(helper, path + "/d:horizontal");
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                Left.SetStyle();
                Right.SetStyle();
                Top.SetStyle();
                Bottom.SetStyle();
            }
        }

        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get 
            {
                return Left.HasValue ||
                    Right.HasValue ||
                    Top.HasValue ||
                    Bottom.HasValue||
                    Vertical.HasValue ||
                    Horizontal.HasValue;
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            Left.Clear();
            Right.Clear();
            Top.Clear();
            Bottom.Clear();
            Vertical.Clear();
            Horizontal.Clear();
        }

        public void BorderAround(ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin, eThemeSchemeColor themeColor=eThemeSchemeColor.Accent1)
        {
            Top.Style = borderStyle;
            Top.Color.SetColor(themeColor);
            Right.Style = borderStyle;
            Right.Color.SetColor(themeColor);
            Bottom.Style = borderStyle;
            Bottom.Color.SetColor(themeColor);
            Left.Style = borderStyle;
            Left.Color.SetColor(themeColor);
        }
        public void BorderAround(ExcelBorderStyle borderStyle, Color color)
        {
            Top.Style = borderStyle;
            Top.Color.SetColor(color);
            Right.Style = borderStyle;
            Right.Color.SetColor(color);
            Bottom.Style = borderStyle;
            Bottom.Color.SetColor(color);
            Left.Style = borderStyle;
            Left.Color.SetColor(color);
        }

        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfBorderBase(_styles, _callback) 
            { 
                Bottom = (ExcelDxfBorderItem)Bottom.Clone(), 
                Top= (ExcelDxfBorderItem)Top.Clone(), 
                Left= (ExcelDxfBorderItem)Left.Clone(), 
                Right= (ExcelDxfBorderItem)Right.Clone(),
                Vertical = (ExcelDxfBorderItem)Vertical.Clone(),
                Horizontal = (ExcelDxfBorderItem)Horizontal.Clone(),
            };
        }
        protected internal override void SetValuesFromXml(XmlHelper helper)
        {
            if (helper.ExistsNode("d:border"))
            {
                Left = GetBorderItem(helper, "d:border/d:left", eStyleClass.BorderLeft);
                Right = GetBorderItem(helper, "d:border/d:right", eStyleClass.BorderLeft);
                Bottom = GetBorderItem(helper, "d:border/d:bottom", eStyleClass.BorderLeft);
                Top = GetBorderItem(helper, "d:border/d:top", eStyleClass.BorderLeft);
                Vertical = GetBorderItem(helper, "d:border/d:vertical", eStyleClass.Border);
                Horizontal = GetBorderItem(helper, "d:border/d:horizontal", eStyleClass.Border);
            }
        }
        private ExcelDxfBorderItem GetBorderItem(XmlHelper helper, string path, eStyleClass styleClass)
        {
            ExcelDxfBorderItem bi = new ExcelDxfBorderItem(_styles, styleClass, _callback);
            var exists = helper.ExistsNode(path);
            if (exists)
            {
                var style = helper.GetXmlNodeString(path + "/@style");
                bi.Style = GetBorderStyleEnum(style);
                bi.Color = GetColor(helper, path + "/d:color", styleClass);
            }
            return bi;
        }
        private static ExcelBorderStyle? GetBorderStyleEnum(string style)
        {
            if (style == "") return null;
            string sInStyle = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
            try
            {
                return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
            }
            catch
            {
                return ExcelBorderStyle.None;
            }

        }

    }
}
