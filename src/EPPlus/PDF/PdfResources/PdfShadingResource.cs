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
                var pah = new PdfAxialShading(objectNumber, GradientFillData, version);
                pah.Coords = GradientFillData.coords;
                return pah;
            }
            return null;
        }
    }
}
