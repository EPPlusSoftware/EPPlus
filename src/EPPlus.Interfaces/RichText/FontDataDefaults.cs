using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal class FontDataDefaults : IFontData
    {
        public string Family { get; set; } = "Archivo Narrow";
        public FontSubFamily SubFamily { get; set; } = FontSubFamily.Regular;
        public double FontSize { get; set; } = 11d;
    }
}
