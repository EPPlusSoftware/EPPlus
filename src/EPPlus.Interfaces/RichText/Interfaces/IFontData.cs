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
        string Family {  get; set; }
        FontSubFamily SubFamily { get; set; }
        double FontSize { get; set; }
    }
}
