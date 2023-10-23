using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IBorderItem
    {
        ExcelBorderStyle Style { get; }
        
        IStyleColor? Color { get; }
    }
}