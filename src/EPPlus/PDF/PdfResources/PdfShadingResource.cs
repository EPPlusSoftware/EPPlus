using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfObjects.PdfShadings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfShadingResource : PdfResource
    {
        internal int shadingObjectNumber;
        internal PdfCellGradientFillData GradientFillData;

        public PdfShadingResource(int labelNumber, PdfCellGradientFillData gradientFillData)
            : base("Sh", labelNumber)
        {
            GradientFillData = gradientFillData;
        }

        public PdfShading GetShadingObject(int objectNumber, int version = 0)
        {
            this.shadingObjectNumber = objectNumber;
            if (GradientFillData != null)
            {
                if (GradientFillData.GradientType == Style.ExcelFillGradientType.Linear)
                {
                    var pas = new PdfAxialShading(objectNumber, GradientFillData, version);
                    pas.Coords = GradientFillData.coords;
                    return pas;
                }
                else if (GradientFillData.GradientType == Style.ExcelFillGradientType.Path)
                {
                    var prs = new PdfRadialShading(objectNumber, GradientFillData, version);
                    prs.Coords = GradientFillData.coords;
                    return prs;
                }
            }
            return null;
        }
    }
}
