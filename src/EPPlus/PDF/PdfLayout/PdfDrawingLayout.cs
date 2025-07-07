using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfDrawingLayout : PdfTransform
    {
        public ExcelDrawing Drawing;
        public PdfDrawingLayout(ExcelDrawing drawing, double x, double y, double width, double height)
            : base(x,y,width,height)
        {
            this.Drawing = drawing;
        }
    }
}
