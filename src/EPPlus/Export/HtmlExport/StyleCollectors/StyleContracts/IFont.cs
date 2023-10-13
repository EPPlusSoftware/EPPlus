using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IFont
    {
        string Name { get; }
        float Size { get; }
        IStyleColor Color { get; }

        bool Bold { get; }
        bool Italic { get; }
        bool Strike { get; }
        ExcelUnderLineType UnderLineType { get; }
    }
}
