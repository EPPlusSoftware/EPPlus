using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FontXml : IFont
    {
        ExcelFont _font;

        IStyleColor _color;

        public FontXml(ExcelFont font) 
        { 
            _font = font;
        }

        public string Name
        { 
            get { return _font.Name; } 
        }

        public float Size 
        { 
            get { return _font.Size; } 
        }

        //public IStyleColor Color
        //{
        //    get { return _color; }
        //}

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
    }
}
