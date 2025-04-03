using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using System;
using System.Drawing;

namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    public enum HFFormattingCodes
    {
        Text,
        PageNumber,
        NumberOfPages,
        SheetName,
        FilePath,
        FileName,
        CurrentDate,
        CurrentTime,
        Image,
    }

    public class HFText
    {
        #region Header Footer text properties
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool Underline { get; set; } = false;
        public bool DoubleUnderline { get; set; } = false;
        public Color Color { get; set; } = Color.Empty;
        public eThemeSchemeColor? Theme { get; set; } = null;
        public double? Tint { get; set; } = null;
        public bool Outline { get; set; } = false;
        public bool Shadow { get; set; } = false;
        public bool Striketrough { get; set; } = false;
        public bool SuperScript { get; set; } = false;
        public bool SubScript { get; set; } = false;
        public string? FontName { get; set; } = null;
        public int? FontSize { get; set; } = null;
        public HFFormattingCodes FormatCode = HFFormattingCodes.Text;
        public string PageNumberSuffix = string.Empty;
        #endregion

        public string Text { get; set; }

        public HFText(){}

        public HFText(string text)
        {
            Text = text;
        }
        public string GetThemeOrColorAsString()
        {
            string c = "";
            if (Theme != null) c = $"&K{((int)Theme).ToString("00")}{(Tint >= 0 ? "+" : "-")}{Math.Abs((double)Tint * 100).ToString("000")}";
            else if (Color != Color.Empty) c = $"&K{Color.R:X2}{Color.G:X2}{Color.B:X2}";
            return c;
        }
    }
}
