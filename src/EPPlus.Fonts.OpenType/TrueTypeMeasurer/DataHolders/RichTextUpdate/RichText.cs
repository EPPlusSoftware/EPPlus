using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders.RichTextUpdate
{
    internal class RichText
    {
        internal string Text { get; private set; }
        internal MeasurementFont Font { get; private set; }
        internal double FontSize { get; private set; }

        public RichText(string text, MeasurementFont font) 
        {
            Text = text;
            Font = font;
            FontSize = font.Size;
        }
    }
}
