using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderItemDxf : IBorderItem
    {
        ExcelDxfBorderItem _border;
        IStyleColor _color;

        public BorderItemDxf(ExcelDxfBorderItem border)
        {
            _border = border;
            _color = new StyleColorDxf(border.Color);
        }

        public ExcelBorderStyle Style
        {
            get
            { return _border.Style != null 
                    ? _border.Style.Value : ExcelBorderStyle.None; }
        }


        public IStyleColor Color { get { return _color; } }
    }
}
