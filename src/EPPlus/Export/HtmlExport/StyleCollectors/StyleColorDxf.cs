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