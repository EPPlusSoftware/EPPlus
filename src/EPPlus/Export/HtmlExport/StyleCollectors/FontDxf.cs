using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FontDxf: IFont
    {
        ExcelDxfFontBase _font;

        IStyleColor _color;

        public FontDxf(ExcelDxfFontBase font)
        {
            _font = font;
            _color = new StyleColorDxf(font.Color);
        }

        //No such property by definition
        public string Name
        {
            //TODO: fix
            get { return null; }
        }

        //No such property by definition
        public float Size
        {
            get { return float.NaN; }
        }

        //Implement IColor.
        public IStyleColor Color
        {
            get { return _color; }
        }

        public bool Bold
        {
            get { return _font.Bold.HasValue ? _font.Bold.Value : false; }
        }

        public bool Italic
        {
            get { return _font.Italic.HasValue ? _font.Italic.Value : false; }
        }

        public bool Strike
        {
            get { return _font.Strike.HasValue ? _font.Strike.Value : false; }
        }

        public ExcelUnderLineType UnderLineType
        {
            get { return _font.Underline.HasValue ? _font.Underline.Value : ExcelUnderLineType.None; }
        }

        public bool HasValue
        {
            get
            {
                return _font.HasValue;
            }
        }
    }
}
