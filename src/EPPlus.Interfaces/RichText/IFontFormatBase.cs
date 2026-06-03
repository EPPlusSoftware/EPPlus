using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    public interface IFontFormatBase
    {
        public string Family { get; set; }
        public FontSubFamily SubFamily { get; set; }
        public float Size { get; set; }

        /// <summary>
        /// Sets the underlying data properties to be equal to the font
        /// </summary>
        /// <param name="font"></param>
        public void SetFont(IFontFormatBase font);

        /// <summary>
        /// Sets the underlying data properties to be equal to the font
        /// </summary>
        /// <param name="font"></param>
        public void SetFont(MeasurementFont font);
    }
}
