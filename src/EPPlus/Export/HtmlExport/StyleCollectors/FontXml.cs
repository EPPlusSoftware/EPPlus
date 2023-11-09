using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FontXml : IFont
    {
        ExcelFontXml _font;

        IStyleColor _color;

        public FontXml(ExcelFontXml font) 
        { 
            _font = font;
            _color = new StyleColorXml(font.Color);
        }

        public string Name
        { 
            get { return _font.Name; } 
        }

        public float Size 
        { 
            get { return _font.Size; } 
        }

        public IStyleColor Color
        {
            get { return _color; }
        }

        public bool Bold
        {
            get { return _font.Bold; }
        }

        public bool Italic
        {
            get { return _font.Italic; }
        }

        public bool Strike
        {
            get { return _font.Strike; }
        }

        public ExcelUnderLineType UnderLineType
        {
            get { return _font.UnderLineType; }
        }

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrEmpty(_font.Id);
            }
        }
    }
}
