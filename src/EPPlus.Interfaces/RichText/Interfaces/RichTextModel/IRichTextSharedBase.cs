using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    public interface IRichTextSharedBase : IRichTextInfoEssential
    {
        /// <summary>
        /// The most essential piece to measure
        /// </summary>
        public string Text { get; set; }
    }
}
