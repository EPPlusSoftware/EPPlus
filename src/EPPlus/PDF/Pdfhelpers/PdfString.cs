using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.Pdfhelpers
{
    internal static class PdfString
    {
        internal static string ToPdfString(this double val)
        {
            return val.ToString(CultureInfo.InvariantCulture);
        }
    }
}
