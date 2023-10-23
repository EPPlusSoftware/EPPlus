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
