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
