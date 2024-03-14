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
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleColorDxf : IStyleColor
    {
        ExcelDxfColor _color;

        public StyleColorDxf(ExcelDxfColor color)
        {
            _color = color;
        }

        public bool Auto { get { return _color.Auto != null ? _color.Auto.Value : false; } }

        public int Indexed { get { return _color.Index != null ? _color.Index.Value : int.MinValue; } }

        public string Rgb 
        { 
            get 
            { 
                return _color.Color != null ? _color.Color.Value.ToArgb().ToString("X") : ""; 
            } 
        }

        public eThemeSchemeColor? Theme { get { return _color.Theme; } }

        public double Tint { get { return _color.Tint != null ? _color.Tint.Value : 0; } }

        public bool Exists { get { return _color.HasValue; } }

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