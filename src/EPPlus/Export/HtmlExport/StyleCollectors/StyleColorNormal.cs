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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleColorNormal : IStyleColor
    {
        ExcelColor _color;

        public StyleColorNormal(ExcelColor color)
        {
            _color = color;
        }

        //TODO: Is this correct?
        public bool Exists { get { return true; } }

        public bool Auto { get { return _color.Auto; } }

        public int Indexed { get { return _color.Indexed; } }

        public double Tint { get { return (double)_color.Tint; } }

        public eThemeSchemeColor? Theme { get { return _color.Theme; } }

        public string Rgb { get { return _color.Rgb; } }

        public bool AreColorEqual(IStyleColor color)
        {
            return StyleColorShared.AreColorEqual(this, color);
        }

        public string GetColor(ExcelTheme theme)
        {
            return StyleColorShared.GetColor(this, theme);
        }
    }
}
