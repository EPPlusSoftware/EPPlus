using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfObjects.PdfPatterns;
using OfficeOpenXml.PDF.PdfObjects.PdfShadings;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfPatternResource : PdfResource
    {
        internal int shadingPatternobjectNumber;
        internal PdfCellFillData CellFillData;

        public PdfPatternResource(int labelNumber, PdfCellFillData cellFillData)
            : base("P", labelNumber)
        {
            CellFillData = cellFillData;
        }

        public PdfPattern GetPatternObject(int objectNumber, int version = 0)
        {
            shadingPatternobjectNumber = objectNumber;
            if (CellFillData.GradientFillData != null)
            {
            }
            else if (CellFillData.PattenStyle != Style.ExcelFillStyle.None && CellFillData.PattenStyle != Style.ExcelFillStyle.Solid)
            {
                var tp = new PdfTilingPattern(objectNumber, version);
                switch (CellFillData.PattenStyle)
                {
                    case Style.ExcelFillStyle.DarkGray:
                        break;
                    case Style.ExcelFillStyle.MediumGray:
                        break;
                    case Style.ExcelFillStyle.LightGray:
                        break;
                    case Style.ExcelFillStyle.Gray125:
                        break;
                    case Style.ExcelFillStyle.Gray0625:
                        break;
                    case Style.ExcelFillStyle.DarkVertical:
                        tp.fill = new PdfPatternDarkVertical(CellFillData.PatternColor, CellFillData.BackgroundColor);
                        tp.BBox = [0, 0, 2, 1];
                        tp.XStep = 2;
                        tp.YStep = 1;
                        break;
                    case Style.ExcelFillStyle.DarkHorizontal:
                        tp.fill = new PdfPatternDarkHorizontal(CellFillData.PatternColor, CellFillData.BackgroundColor);
                        tp.BBox = [0, 0, 1, 2];
                        tp.XStep = 1;
                        tp.YStep = 2;
                        break;
                    case Style.ExcelFillStyle.DarkDown:
                        break;
                    case Style.ExcelFillStyle.DarkUp:
                        break;
                    case Style.ExcelFillStyle.DarkGrid:
                        break;
                    case Style.ExcelFillStyle.DarkTrellis:
                        break;
                    case Style.ExcelFillStyle.LightVertical:
                        break;
                    case Style.ExcelFillStyle.LightHorizontal:
                        break;
                    case Style.ExcelFillStyle.LightDown:
                        break;
                    case Style.ExcelFillStyle.LightUp:
                        break;
                    case Style.ExcelFillStyle.LightGrid:
                        break;
                    case Style.ExcelFillStyle.LightTrellis:
                        break;
                }
                return tp;
            }
            return null;
        }

        public PdfShading GetShadingObject(int objectNumber, int version = 0)
        {
            if (CellFillData.GradientFillData != null)
            {
                return new PdfAxialShading(objectNumber, CellFillData.GradientFillData, version);
            }
            return null;
        }

        public PdfShadingPattern GetShadingPatternObject(int objectNumber, int shadingObjectNumber, int version = 0)
        {
            shadingPatternobjectNumber = objectNumber;
            var shadingPattern = new PdfShadingPattern(objectNumber, shadingObjectNumber);
            shadingPattern.Matrix = CellFillData.GradientFillData.matrix;
            return shadingPattern;
        }
    }
}
