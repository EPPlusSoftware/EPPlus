using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderXml : IBorder
    {
        BorderItemXml _top;
        BorderItemXml _bottom;
        BorderItemXml _left;
        BorderItemXml _right;

        internal BorderXml(ExcelBorderXml border)
        {
            _top = new BorderItemXml(border.Top);
            _bottom = new BorderItemXml(border.Bottom);
            _left = new BorderItemXml(border.Left);
            _right = new BorderItemXml(border.Right);
        }

        public IBorderItem Top { get { return _top; } }

        public IBorderItem Bottom { get { return _bottom; } }

        public IBorderItem Left { get { return _left; } }

        public IBorderItem Right { get { return _right; } }
    }
}
