using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// Collects size and font family names
    /// </summary>
    internal interface IFontData
    {
        string Family {  get; set; }
        FontSubFamily SubFamily { get; set; }
        double FontSize { get; set; }
    }
}
