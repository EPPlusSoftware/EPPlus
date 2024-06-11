using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Drawing;

namespace OfficeOpenXml.SystemDrawing.Text
{
    public class SystemDrawingTextMeasurer : ITextMeasurer, IDisposable
    {
        public SystemDrawingTextMeasurer()
        {
            _stringFormat = StringFormat.GenericDefault;
            _bmp = new Bitmap(1, 1);
            _graphics = System.Drawing.Graphics.FromImage(_bmp);
        }

        private readonly StringFormat _stringFormat;
        private readonly Graphics _graphics;
        private readonly Bitmap _bmp;
        private bool disposedValue;

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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~SystemDrawingTextMeasurer()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            _bmp.Dispose();
            _graphics.Dispose();
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            float dpiCorrectX, dpiCorrectY;
            try
            {
                //Check for missing GDI+, then use WPF istead.
                _graphics.PageUnit = GraphicsUnit.Pixel;
                dpiCorrectX = 96 / _graphics.DpiX;
                dpiCorrectY = 96 / _graphics.DpiY;
            }
            catch
            {
                return TextMeasurement.Empty;
            }
            var style = ToFontStyle(font.Style);
            var dFont = new Font(font.FontFamily, font.Size, style);
            var size = _graphics.MeasureString(text, dFont, 10000, _stringFormat);
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
