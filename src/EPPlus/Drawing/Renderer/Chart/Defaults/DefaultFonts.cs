using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class DefaultFonts2007Style
    {
        internal static readonly MeasurementFont Base = new MeasurementFont() { Size = 10 };
        internal static readonly MeasurementFont Title = new MeasurementFont() { Size = 18, Style=MeasurementFontStyles.Bold };
        internal static readonly MeasurementFont AxisTitle = new MeasurementFont() { Size = 10, Style = MeasurementFontStyles.Bold };
        internal static readonly MeasurementFont Axis = new MeasurementFont() { Size = 10 };
        internal static readonly MeasurementFont Legend = new MeasurementFont() { Size = 10 };
        internal static readonly MeasurementFont DataTable = new MeasurementFont() { Size = 10 };
        internal static readonly MeasurementFont DataLabels = new MeasurementFont() { Size = 10 };
        internal static readonly MeasurementFont Shape = new MeasurementFont() { Size = 11 };
    }
}
