using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal class RichTextBase : IRichText
    {
        internal bool FirstInParagraph { get; private set; }
        internal RichTextBase(string text, bool firstInParagraph, string fontFamily = "") 
        {
            Text = text;
            FirstInParagraph = firstInParagraph;
            if (string.IsNullOrEmpty(fontFamily) == false)
            {
                RichTextOptions.FontFamily = fontFamily;
            }
        }

        public IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();
        public string Text { get; set; }
    }
}
