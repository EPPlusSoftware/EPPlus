using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FontDxf: IFont
    {
        ExcelDxfStyleLimitedFont _font;

        IStyleColor _color;

        public FontDxf(ExcelDxfStyleLimitedFont font)
        {
            _font = font;
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
        //public IStyleColor Color
        //{
        //    get { return _color; }
        //}

        public bool Bold
        {
            get { return _font.Font.Bold.HasValue ? _font.Font.Bold.Value : false; }
        }

        public bool Italic
        {
            get { return _font.Font.Italic.HasValue ? _font.Font.Italic.Value : false; }
        }

        public bool Strike
        {
            get { return _font.Font.Strike.HasValue ? _font.Font.Strike.Value : false; }
        }

        public ExcelUnderLineType UnderLineType
        {
            get { return _font.Font.Underline.HasValue ? _font.Font.Underline.Value : ExcelUnderLineType.None; }
        }
    }
}
