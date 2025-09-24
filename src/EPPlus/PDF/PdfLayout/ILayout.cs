using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal interface ILayout
    {
        public void ConvertCoordinates(PdfPageSettings pageSettings);
    }
}
