using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Drawing;

namespace OfficeOpenXml.SystemDrawing.Text
{
    public class SystemDrawingTextMeasurer : ITextMeasurer
    {
        public SystemDrawingTextMeasurer()
        {
            _stringFormat = StringFormat.GenericDefault;
        }

        private readonly StringFormat _stringFormat;
        private FontStyle ToFontStyle(MeasurementFontStyles fontStyle)
        {
            switch (fontStyle)
            {
                case MeasurementFontStyles.Bold | MeasurementFontStyles.Italic:
                    return FontStyle.Bold | FontStyle.Italic;
                case MeasurementFontStyles.Regular:
                    return FontStyle.Regular;
                case MeasurementFontStyles.Bold:
                    return FontStyle.Bold;
                case MeasurementFontStyles.Italic:
                    return FontStyle.Italic;
                default:
                    return FontStyle.Regular;
            }
        }        
        public TextMeasurement MeasureText(string text, MeasurementFont font)
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
            return new TextMeasurement(size.Width * dpiCorrectX, size.Height * dpiCorrectY);
        }
        bool? _validForEnvironment=null;
        public bool ValidForEnvironment()
        {
            if(_validForEnvironment.HasValue==false)
            {
                try
                {
                    var g=Graphics.FromHwnd(IntPtr.Zero);
                    g.MeasureString("d",new Font("Calibri", 11, FontStyle.Regular));
                    _validForEnvironment = true;
                }
                catch 
                { 
                    _validForEnvironment = false;
                }
            }
            return _validForEnvironment.Value;
        }
    }
}
