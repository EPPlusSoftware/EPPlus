using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Rich text info that is Essential for measuring correctly
    /// Interface for pdf/svg/future richtext implementations to unify richtext styling
    /// </summary>
    public interface IRichTextInfoEssential : IFontData
    {
        ///// <summary>
        ///// The text within the rich text
        ///// The most essential part to measure
        ///// </summary>
        //string Text { get; set; }

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
