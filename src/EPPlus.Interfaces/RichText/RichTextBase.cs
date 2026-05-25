using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class RichTextBase : IRichText
    {
        internal bool FirstInParagraph { get; private set; }
        public RichTextBase(string text, bool firstInParagraph, string fontFamily = "") 
        {
            Text = text;
            FirstInParagraph = firstInParagraph;
            if (string.IsNullOrEmpty(fontFamily) == false)
            {
                Info.FamilyName = fontFamily;
            }
        }

        public IRichTextInfoSimple Info { get; set; } = new RichTextDefaults();
        public string Text { get; set; }
    }
}
