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
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A color in a differential formatting record
    /// </summary>
    public class ExcelDxfColor : DxfStyleBase
    {
        eStyleClass _styleClass;
        internal ExcelDxfColor(ExcelStyles styles, eStyleClass styleClass, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
            _styleClass = styleClass;
        }
        eThemeSchemeColor? _theme=null;
        /// <summary>
        /// Gets or sets a theme color
        /// </summary>
        public eThemeSchemeColor? Theme 
        { 
            get
            {
                return _theme;
            }
            set
            {
                _theme = value;
                _callback?.Invoke(_styleClass, eStyleProperty.Theme, value);
            }
        }
        int? _index;
        /// <summary>
        /// Gets or sets an indexed color
        /// </summary>
        public int? Index
        {
            get
            {
                return _index;
            }
            set
            {
                _index = value;
                _callback?.Invoke(_styleClass, eStyleProperty.IndexedColor, value);
            }
        }
        bool? _auto;
        /// <summary>
        /// Gets or sets the color to automatic
        /// </summary>
        public bool? Auto
        {
            get
            {
                return _auto;
            }
            set
            {
                _auto = value;
                _callback?.Invoke(_styleClass, eStyleProperty.AutoColor, value);
            }
        }
        double? _tint;
        /// <summary>
        /// Gets or sets the Tint value for the color
        /// </summary>
        public double? Tint
        {
            get
            {
                return _tint;
            }
            set
            {
                _tint = value;
                _callback?.Invoke(_styleClass, eStyleProperty.Tint, value);
            }
        }
        Color? _color;
        /// <summary>
        /// Sets the color.
        /// </summary>
        public Color? Color 
        {
            get
            {
                return _color;
            }
            set
            {
                _color = value;
                _callback?.Invoke(_styleClass, eStyleProperty.Color, value);
            }
        }
        /// <summary>
        /// The Id
        /// </summary>
        internal override string Id
        {
            get { return GetAsString(Theme) + "|" + GetAsString(Index) + "|" + GetAsString(Auto) + "|" + GetAsString(Tint) + "|" + GetAsString(Color==null ? "" : Color.Value.ToArgb().ToString("x")); }
        }
        internal static string GetEmptyId()
        {
			return "||||";
		}
        /// <summary>
        /// Set the color of the drawing based on an RGB color. This method will remove any previous 
        /// <see cref="eThemeSchemeColor">ThemeSchemeColor</see>, <see cref="ExcelIndexedColor">IndexedColor</see> 
        /// or Automatic color used.
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
        /// Set the color of the drawing based on an <see cref="eThemeSchemeColor"/> color. 
        /// This method will remove any previous <see cref="System.Drawing.Color"/>, 
        /// <see cref="ExcelIndexedColor">IndexedColor</see> or Automatic color used.
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
        /// Set the color of the drawing based on an <see cref="ExcelIndexedColor"/> color. 
        /// This method will remove any previous <see cref="System.Drawing.Color">Color</see>, 
        /// <see cref="eThemeSchemeColor">ThemeSchemeColor</see> or Automatic color used.
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
        /// Set the color of the object
        /// </summary>
        /// <param name="alpha">Alpha component value</param>
        /// <param name="red">Red component value</param>
        /// <param name="green">Green component value</param>
        /// <param name="blue">Blue component value</param>
        public void SetColor(int alpha, int red, int green, int blue)
        {
            if (alpha < 0 || red < 0 || green < 0 || blue < 0 ||
               alpha > 255 || red > 255 || green > 255 || blue > 255)
            {
                throw (new ArgumentException("Argument range must be from 0 to 255"));
            }
            SetColor(System.Drawing.Color.FromArgb(alpha, red, green, blue));
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                _callback.Invoke(_styleClass, eStyleProperty.Color, _color);
                _callback.Invoke(_styleClass, eStyleProperty.Theme, _theme);
                _callback.Invoke(_styleClass, eStyleProperty.IndexedColor, _index);
                _callback.Invoke(_styleClass, eStyleProperty.AutoColor, _auto);
                _callback.Invoke(_styleClass, eStyleProperty.Tint, _tint);
            }
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        internal override DxfStyleBase Clone()
        {
            return new ExcelDxfColor(_styles, _styleClass, _callback) { Theme = Theme, Index = Index, Color = Color, Auto = Auto, Tint = Tint };
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
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
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            Theme = null;
            Index = null;
            Auto = null;
            Tint = null;
            Color = null;
        }
        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            throw new NotImplementedException();
        }

        internal Color GetColorAsColor(bool whiteAsDefault = false)
        {
            if (Index != null)
            {
                return _styles.GetIndexedColor(Index.Value);
            }
            else if (Color != null)
            {
                return Color.Value;
            }
            else if (Theme.HasValue)
            {
                var themeColor = _styles._wb.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum(Theme.Value);
                return Utils.ColorConverter.GetThemeColor(themeColor);
            }
            else if (Auto.HasValue)
            {
                var themeColor = _styles._wb.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum(eThemeSchemeColor.Background1);
                return Utils.ColorConverter.GetThemeColor(themeColor);
            }
            else if(whiteAsDefault)
            {
                return System.Drawing.Color.White;
            }

            return System.Drawing.Color.Empty;
        }

        /// <summary>
        /// Return the RGB value as a string for the color object that uses the Indexed or Tint property
        /// </summary>
        /// <returns>The RGB color starting with a #FF (alpha)</returns>
        internal string LookupColor()
        {
            if (Index >= 0 && Index < _styles.IndexedColors.Length)
            {
                return _styles.IndexedColors[Index.Value];
            }
            else if (Color != null)
            {
                return "#" + Color.Value.ToArgb().ToString("x8").Substring(2);
            }
            else if (Theme.HasValue)
            {
                return GetThemeColor(Theme.Value, Convert.ToDouble(Tint));
            }
            else if (Auto.Value)
            {
                return GetThemeColor(eThemeSchemeColor.Background1, Convert.ToDouble(Tint));
            }
            else
            {
                string c = "F";
                if(Tint != null)
                {
                    c = ((int)(Math.Round((Tint.Value + 1) * 128))).ToString("X");
                }
                return "#FF" + c + c + c;
            }
        }

        private string GetThemeColor(eThemeSchemeColor theme, double tint)
        {
            var themeColor = _styles._wb.ThemeManager.GetOrCreateTheme().ColorScheme.GetColorByEnum(theme);
            var color = Utils.ColorConverter.GetThemeColor(themeColor);
            if (tint != 0)
            {
                color = Utils.ColorConverter.ApplyTint(color, tint);
            }

            return "#" + color.ToArgb().ToString("X");
        }
    }
}
