using OfficeOpenXml.Interfaces.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.System.Drawing.Text
{
    public class SystemDrawingTextMeasurer : ITextMeasurer
    {
        public SystemDrawingTextMeasurer()
        {
            _stringFormat = StringFormat.GenericDefault;
        }

        private readonly StringFormat _stringFormat;
        private FontStyle ToFontStyle(FontStyles fontStyle)
        {
            switch(fontStyle)
            {
                case FontStyles.Bold | FontStyles.Italic:
                    return FontStyle.Bold | FontStyle.Italic;
                case FontStyles.Regular:
                    return FontStyle.Regular;
                case FontStyles.Bold:
                    return FontStyle.Bold;
                case FontStyles.Italic:
                    return FontStyle.Italic;
                default:
                    return FontStyle.Regular;
            }
        }
        public TextMeasurement MeasureText(string text, ExcelFont font)
        {
            Bitmap b;
            Graphics g;
            float dpiCorrectX, dpiCorrectY;
            try
            {
                //Check for missing GDI+, then use WPF istead.
                b = new Bitmap(1, 1);
                g = Graphics.FromImage(b);
                g.PageUnit = GraphicsUnit.Pixel;
                dpiCorrectX = 96 / g.DpiX;
                dpiCorrectY = 96 / g.DpiY;
            }
            catch
            {
                return TextMeasurement.Empty;
            }
            var style = ToFontStyle(font.Style);
            var dFont = new Font(font.FontFamily, font.Size, style);
            var size = g.MeasureString(text, dFont, 10000, _stringFormat);
            return new TextMeasurement(size.Width, size.Height);
        }
    }
}
