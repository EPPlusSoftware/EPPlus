using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    public interface ITextFragment
    {
        string Text { get; set; }

        IRichText RichTextOptions { get; set; }
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }
    }
}
