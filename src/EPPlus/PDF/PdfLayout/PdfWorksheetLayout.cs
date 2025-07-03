using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.PDF.Pdfhelpers;
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
            double x = 0d;
            double y = 0d;
            List<string> checkedMergedCells = new List<string>();
            for(int i = 1; i<= ws.Dimension._toRow; i++)
            {
                var height = ws.Row(i).Height;
                for (int j = 1; j <= ws.Dimension._toCol; j++)
                {
                    var cell = ws.Cells[i, j];
                    bool isMerged = false;
                    if (cell.Merge)
                    {
                        isMerged = true;
                        if (!checkedMergedCells.Contains(ws.MergedCells[i, j]))
                        {
                            double mcHeight = 0;
                            double mcWidth = 0;
                            ExcelAddressBase address = new ExcelAddressBase(ws.MergedCells[i, j]);
                            for (int k = address._fromRow; k <= address._toRow; k++)
                            {
                                mcHeight += ws.Row(k).Height;
                            }
                            for (int l = address._fromCol; l <= address._toCol; l++)
                            {
                                mcWidth += (ws.Column(l).Width);
                            }
                            objects.Add(new PdfCellLayout(ConvertCellValueToTransform(ws.Cells[address._fromRow, address._fromCol].Value), x, y, mcWidth, mcHeight));
                            checkedMergedCells.Add(ws.MergedCells[i, j]);
                        }
                    }
                    var width = PdfUnits.ExcelColumnWidthToPoints(ws.Column(j).Width);
                    objects.Add(new PdfCellLayout(isMerged ? null : ConvertCellValueToTransform(cell.Value), x, y, width, height));
                    x+= width;
                }
                y += height;
                x = 0d;
            }
        }


        private PdfTransform ConvertCellValueToTransform(object value)
        {
            if (value is string)
            {

            }
        }

    }
}
