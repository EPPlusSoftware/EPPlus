using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfObjects.PdfPatterns;
using OfficeOpenXml.PDF.PdfObjects.PdfShadings;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfPatternResource : PdfResource
    {
        internal int shadingPatternobjectNumber;
        internal PdfCellGradientFillData GradientFillData;

        public PdfPatternResource(int labelNumber, PdfCellGradientFillData gradientFillData)
            : base("P", labelNumber)
        {
            GradientFillData = gradientFillData;
        }

        public PdfShading GetShadingObject(int objectNumber, int version = 0)
        {
            if (GradientFillData != null)
            {
                return new PdfAxialShading(objectNumber, GradientFillData, version);
            }
            return null;
        }

        public PdfShadingPattern GetShadingPatternObject(int objectNumber, int shadingObjectNumber, int version = 0)
        {
            shadingPatternobjectNumber = objectNumber;
            var shadingPattern = new PdfShadingPattern(objectNumber, shadingObjectNumber);
            shadingPattern.Matrix = GradientFillData.matrix;
            return shadingPattern;
        }
    }
}
