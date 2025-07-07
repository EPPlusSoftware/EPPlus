using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellContentLayout : PdfTransform
    {
        internal object Value;

        public PdfCellContentLayout(object value, double x, double y, double width, double height)
            : base(x, y, width, height)
        {
            Value = value;
        }


        //implement text handling here like length measure and where to to do new line and cut off.
    }
}
