using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    public interface IRichTextFormatBase : IFontFormatBase
    {
        /// <summary>
        /// The text within the rich text
        /// The most essential part to measure
        /// </summary>
        string Text { get; set; }

        /// <summary>
        /// This MUST interact with font data subfamily
        /// </summary>
        bool Italic { get; set; }
        /// <summary>
        /// This MUST interact with font data subfamily
        /// </summary>
        bool Bold { get; set; }
    }
}
