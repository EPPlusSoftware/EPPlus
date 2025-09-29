using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    public interface ITextMeasurerWrap : ITextMeasurer
    {
        public List<string> MeasureAndWrapText(string text, MeasurementFont font, double MaxWidthInPixels);

        public double GetSingleLineSpacing();

        /// <summary>
        /// Get only the ascendant, i.e. space between top of the text font space and its baseline
        /// </summary>
        /// <returns></returns>
        public double GetBaseLine();
    }
}
