using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderItemXml : IBorderItem

    {
        ExcelBorderItemXml _border;
        IStyleColor _color;

        public BorderItemXml(ExcelBorderItemXml border) 
        {
            _border = border;
            _color = new StyleColorXml(border.Color);
        }

        public ExcelBorderStyle Style { get { return _border.Style; } }

        public IStyleColor Color { get { return _color; } }
    }
}
