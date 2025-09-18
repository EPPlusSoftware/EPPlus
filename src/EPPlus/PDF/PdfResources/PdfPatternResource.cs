using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfPatternResource : PdfResource
    {
        internal int objectNumber;
        internal PdfCellGradientFillData GradientData;

        public PdfPatternResource(int labelNumber)
            : base("P", labelNumber)
        {

        }

    }
}
