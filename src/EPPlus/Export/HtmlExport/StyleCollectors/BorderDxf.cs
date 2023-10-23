

using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderDxf : IBorder
    {
        BorderItemDxf _top;
        BorderItemDxf _bottom;
        BorderItemDxf _left;
        BorderItemDxf _right;

        internal BorderDxf(ExcelDxfBorderBase border)
        {
            _top = new BorderItemDxf(border.Top);
            _bottom = new BorderItemDxf(border.Bottom);
            _left = new BorderItemDxf(border.Left);
            _right = new BorderItemDxf(border.Right);
        }

        public IBorderItem Top { get { return _top; } }

        public IBorderItem Bottom { get { return _bottom; } }

        public IBorderItem Left { get { return _left; } }

        public IBorderItem Right { get { return _right; } }
    }
}
