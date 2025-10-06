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
                switch (CellFillData.PattenStyle)
                {
                    case Style.ExcelFillStyle.DarkGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 2], 4, 2, version);
                    case Style.ExcelFillStyle.MediumGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternMediumGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 2], 2, 2, version);
                    case Style.ExcelFillStyle.LightGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 2], 4, 2, version);
                    case Style.ExcelFillStyle.Gray125:
                        return new PdfTilingPattern(objectNumber, new PdfPatternGray125(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.Gray0625:
                        return new PdfTilingPattern(objectNumber, new PdfPatternGray0625(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 8, 4], 8, 4, version);
                    case Style.ExcelFillStyle.DarkVertical:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkVertical(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 1], 2, 1, version);
                    case Style.ExcelFillStyle.DarkHorizontal:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkHorizontal(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 1, 2], 1, 2, version);
                    case Style.ExcelFillStyle.DarkDown:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkDown(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 2], 2, 2, version);
                    case Style.ExcelFillStyle.DarkUp:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkUp(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.DarkGrid:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkGrid(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 2], 2, 2, version);
                    case Style.ExcelFillStyle.DarkTrellis:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkTrellis(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.LightVertical:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightVertical(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 1, 0.5d], 1, 0.5, version);
                    case Style.ExcelFillStyle.LightHorizontal:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightHorizontal(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 0.5d, 1], 0.5d, 1, version);
                    case Style.ExcelFillStyle.LightDown:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightDown(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.LightUp:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightUp(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.LightGrid:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightGrid(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case Style.ExcelFillStyle.LightTrellis:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightTrellis(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                }
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
