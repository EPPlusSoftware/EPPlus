﻿using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleColorXml : IStyleColor
    {
        ExcelColorXml _color;

        public StyleColorXml(ExcelColorXml color) 
        {
            _color = color;
        }

        public bool Exists { get { return _color.Exists; } }

        public bool Auto { get { return _color.Auto; } }

        public int Indexed { get { return _color.Indexed; } }

        public double Tint { get { return (double)_color.Tint; } }

        public eThemeSchemeColor? Theme { get { return _color.Theme; } }

        public string Rgb { get { return _color.Rgb; } }

        public bool AreColorEqual(IStyleColor color)
        {
            return StyleColorShared.AreColorEqual(this, color);
        }

        public string GetColor(ExcelTheme theme)
        {
            return StyleColorShared.GetColor(this, theme);
        }
    }
}
