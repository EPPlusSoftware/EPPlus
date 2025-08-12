using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{

    internal class PdfCellContentLayout : PdfTransform
    {



        public PdfCellContentLayout(ExcelRangeBase Cell, double x, double y, double width, double height)
            : base(x, y, width, height)
        {

        }


        //implement text handling here like length measure and where to to do new line and cut off.
    }
}
