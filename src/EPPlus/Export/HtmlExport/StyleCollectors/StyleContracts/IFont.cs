using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IFont
    {
        string Name { get; }
        float Size { get; }
        IStyleColor Color { get; }
        bool HasValue { get; }
        bool Bold { get; }
        bool Italic { get; }
        bool Strike { get; }
        ExcelUnderLineType UnderLineType { get; }
    }
}
