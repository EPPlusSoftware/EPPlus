using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style.HeaderFooterTextFormat
{
    public class HFText
    {
        #region Header Footer text properties
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool Underline { get; set; } = false;
        public bool DoubleUnderline { get; set; } = false;
        public Color Color { get; set; } = Color.Empty;
        public eThemeSchemeColor? theme { get; set; } = null;
        public double? tint { get; set; } = null;
        public bool Outline { get; set; } = false;
        public bool Shadow { get; set; } = false;
        public bool Striketrough { get; set; } = false;
        public bool SuperScript { get; set; } = false;
        public bool SubScript { get; set; } = false;
        public string? FontName { get; set; } = null;
        public int? FontSize { get; set; } = null;
        #endregion

        public string Text { get; set; }

        public HFText(){}

        public HFText(string text)
        {
            Text = text;
        }
    }
}
