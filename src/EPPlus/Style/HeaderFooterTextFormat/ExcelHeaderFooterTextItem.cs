using OfficeOpenXml.Drawing;
using System;
using System.Drawing;

namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    /// <summary>
    /// Enum for inserting format codes into an ExcelHeaderFooterTextCollection.
    /// </summary>
    public enum ExcelHeaderFooterFormattingCodes
    {
        /// <summary>
        /// Indicates that the format is a text.
        /// </summary>
        Text,
        /// <summary>
        /// Inserts page number.
        /// </summary>
        PageNumber,
        /// <summary>
        /// Inserts the number of pages.
        /// </summary>
        NumberOfPages,
        /// <summary>
        /// Inserts sheet name.
        /// </summary>
        SheetName,
        /// <summary>
        /// Inserts file path.
        /// </summary>
        FilePath,
        /// <summary>
        /// Inserts file name.
        /// </summary>
        FileName,
        /// <summary>
        /// Inserts current date.
        /// </summary>
        CurrentDate,
        /// <summary>
        /// Inserts current time.
        /// </summary>
        CurrentTime,
        /// <summary>
        /// Indicates an image is inserted.
        /// </summary>
        Image,
    }

    /// <summary>
    /// An ExcelHeaderFooterTextItem object used for the ExcelHeaderFooterTextCollection
    /// </summary>
    public class ExcelHeaderFooterTextItem
    {
        #region Header Footer text properties
        /// <summary>
        /// If text is bold.
        /// </summary>
        public bool Bold { get; set; } = false;
        /// <summary>
        /// If text is italic.
        /// </summary>
        public bool Italic { get; set; } = false;
        /// <summary>
        /// If text is underlined.
        /// </summary>
        public bool Underline { get; set; } = false;
        /// <summary>
        /// If text is double underlined.
        /// </summary>
        public bool DoubleUnderline { get; set; } = false;
        /// <summary>
        /// Color of text.
        /// </summary>
        public Color Color { get; set; } = Color.Empty;
        /// <summary>
        /// The text theme.
        /// </summary>
        public eThemeSchemeColor? Theme { get; set; } = null;
        /// <summary>
        /// Tint of the text theme.
        /// </summary>
        public double? Tint { get; set; } = null;
        /// <summary>
        /// If text is outlined.
        /// </summary>
        public bool Outline { get; set; } = false;
        /// <summary>
        /// If text has a shadow.
        /// </summary>
        public bool Shadow { get; set; } = false;
        /// <summary>
        /// If text has a strikethrough.
        /// </summary>
        public bool Striketrough { get; set; } = false;
        /// <summary>
        /// If text is superscript.
        /// </summary>
        public bool SuperScript { get; set; } = false;
        /// <summary>
        /// If text is subscript.
        /// </summary>
        public bool SubScript { get; set; } = false;
        /// <summary>
        /// The text font name.
        /// </summary>
        public string FontName { get; set; } = string.Empty;
        /// <summary>
        /// The text size.
        /// </summary>
        public int? FontSize { get; set; } = null;
        /// <summary>
        /// The text format code. Used for inserting dates, page numbers and more.
        /// </summary>
        public ExcelHeaderFooterFormattingCodes FormatCode = ExcelHeaderFooterFormattingCodes.Text;
        /// <summary>
        /// Suffix for page number.
        /// </summary>
        public string PageNumberSuffix = string.Empty;
        #endregion
        /// <summary>
        /// The text.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Creates an empty ExcelHeaderFooterTextItem object.
        /// </summary>
        public ExcelHeaderFooterTextItem(){}

        /// <summary>
        /// Creates an ExcelHeaderFooterTextItem object with a text.
        /// </summary>
        /// <param name="text">The objects text</param>
        public ExcelHeaderFooterTextItem(string text)
        {
            Text = text;
        }

        /// <summary>
        /// Get the theme or color as a string in format &amp;KTTSNNN for theme and &amp;KRRGGBB for color.
        /// </summary>
        /// <returns></returns>
        public string GetThemeOrColorAsString()
        {
            string c = "";
            if (Theme != null) c = $"&K{((int)Theme).ToString("00")}{(Tint >= 0 ? "+" : "-")}{Math.Abs((double)Tint * 100).ToString("000")}";
            else if (Color != Color.Empty) c = $"&K{Color.R:X2}{Color.G:X2}{Color.B:X2}";
            return c;
        }
    }
}
