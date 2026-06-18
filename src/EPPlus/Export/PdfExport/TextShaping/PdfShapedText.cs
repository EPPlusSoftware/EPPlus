/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.TextShaping
{
    internal struct PdfShapedText
    {
        public IFontProvider FontProvider;
        //sometimes a font can have other fonts for certain characters. Key as the glyph id, Value is the font label.
        //public Dictionary<byte, string> FontIDLabel;
        public Dictionary<byte, string> FontIdMap;
        public List<OpenTypeFont> UsedFonts;
        public ShapedText ShapedText;
    }
}
