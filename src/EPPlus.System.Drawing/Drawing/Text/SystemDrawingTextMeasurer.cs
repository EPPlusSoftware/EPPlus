/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           EPPlus Interfaces 1.0
 *************************************************************************************************/
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

        public float GetScalingFactorRowHeight(MeasurementFont font)
        {
            if(font == null || string.IsNullOrEmpty(font.FontFamily))
            {
                return 1f;
            }
            switch(font.FontFamily)
            {
                case "Arial":
                    return 1.02f;
                case "Times New Roman":
                    return 1.15f;
                case "Liberation Serif":
                    return 1.2f;
                case "Verdana":
                    return 1f;
                case "Century Gothic":
                    return 0.95f;
                case "Courier New":
                    return 1.1f;
                case "Arial Black":
                    return 1.02f;
                case "Corbel":
                    return 1.05f;
                case "Rockwell":
                    return 0.97f;
                case "Tw Cen MT":
                    return 1.12f;
                case "Tw Cen MT Condensed":
                    return 1.11f;
                default:
                    return 1f;
            }
        }
    }
}
