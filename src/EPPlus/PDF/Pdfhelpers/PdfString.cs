using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.Pdfhelpers
{
    internal static class PdfString
    {
        internal static string Convert(double val)
        {
            return val.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
