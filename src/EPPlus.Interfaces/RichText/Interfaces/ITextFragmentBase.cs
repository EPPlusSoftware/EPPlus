using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Most basic possible definition of a textFragment/run
    /// </summary>
    public interface ITextFragmentBase
    {
        public string Text { get; set; }
        public IRichTextInfoEssential RichText { get; }
    }
}
