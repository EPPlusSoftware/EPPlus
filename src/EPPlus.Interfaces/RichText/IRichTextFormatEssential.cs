using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// The most basic rich text format
    /// The only properties that belong in this class are those that are absolutely neccesary for Measuring the text correctly
    /// </summary>
    public interface IRichTextFormatEssential : IFontFormatBase
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
