using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal class RichTextBase : IRichText
    {
        internal RichTextBase(string text, string fontFamily = "") 
        {
            Text = text;
            if (string.IsNullOrEmpty(fontFamily) == false)
            {
                FontData.Family = fontFamily;
            }
        }

        public IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();
        public IFontData FontData { get; set; } = new FontDataBasic();
        public string Text { get; set; }
    }
}
