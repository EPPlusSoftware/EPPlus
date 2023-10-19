using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IStyleColor
    {
        bool Exists { get; }
        bool Auto { get; }
        string Rgb { get; }
        int Indexed { get; }
        double Tint { get; }
        eThemeSchemeColor? Theme { get; }
        bool AreColorEqual(IStyleColor color);
    }
}
