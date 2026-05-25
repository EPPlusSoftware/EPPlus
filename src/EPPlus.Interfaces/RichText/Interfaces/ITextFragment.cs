using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// Additional input and measurement data for shaper/layoutengine
    /// </summary>
    public interface ITextFragment : ITextFragmentBase
    {
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }
    }
}
