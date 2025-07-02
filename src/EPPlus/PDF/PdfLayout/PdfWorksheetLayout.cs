using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfWorksheetLayout : PdfTransform
    {
        internal ExcelWorksheet ws;

        internal List<PdfTransform> objects;

        public PdfWorksheetLayout(ExcelWorksheet worksheet)
        {
            this.ws = worksheet;

            for(int i = 1; i<= ws.Dimension._toRow; i++)
            {
                for (int j = 1; j <= ws.Dimension._toCol; j++)
                {
                    var cell = ws.Cells[i, j];
                }
            }

        }
    }
}
