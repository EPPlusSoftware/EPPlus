using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class FontDataDefaults : IFontData
    {
        public string FontFamily { get; set; } = "Archivo Narrow";
        public FontSubFamily SubFamily { get; set; } = FontSubFamily.Regular;
        public float Size { get; set; } = 11f;

        public void SetFont(IFontData font)
        {
            FontFamily = font.FontFamily;
            Size = font.Size;
            SubFamily = font.SubFamily;
        }
    }
}
