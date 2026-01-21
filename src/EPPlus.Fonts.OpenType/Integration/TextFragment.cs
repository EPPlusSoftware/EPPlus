using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Represents a text fragment with specific font properties.
    /// </summary>
    public class TextFragment
    {
        public string Text { get; set; }
        public MeasurementFont Font { get; set; }
        public ShapingOptions Options { get; set; }
    }
}
