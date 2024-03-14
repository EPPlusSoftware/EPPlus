/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

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
