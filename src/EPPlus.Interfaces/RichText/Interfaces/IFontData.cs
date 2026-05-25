using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Collects size and font family names
    /// </summary>
    public interface IFontData
    {
        string FontFamily {  get; set; }
        FontSubFamily SubFamily { get; set; }
        float Size { get; set; }


        /// <summary>
        /// Sets the underlying data properties to be equal to the font
        /// </summary>
        /// <param name="font"></param>
        public void SetFont(IFontData font);
    }
}
