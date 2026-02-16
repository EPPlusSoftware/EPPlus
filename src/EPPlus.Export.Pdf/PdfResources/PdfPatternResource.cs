/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfObjects.PdfPatterns;
using OfficeOpenXml.Style;

namespace EPPlus.Export.Pdf.PdfResources
{
    internal class PdfPatternResource : PdfResource
    {
        internal int objectNumber;
        internal PdfCellFillData CellFillData;

        public PdfPatternResource(int labelNumber, PdfCellFillData cellFillData)
            : base("P", labelNumber)
        {
            CellFillData = cellFillData;
        }

        public PdfPattern GetPatternObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
            if (CellFillData.PatternStyle != ExcelFillStyle.None && CellFillData.PatternStyle != ExcelFillStyle.Solid)
            {
                switch (CellFillData.PatternStyle)
                {
                    case ExcelFillStyle.DarkGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 2], 4, 2, version);
                    case ExcelFillStyle.MediumGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternMediumGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 2], 2, 2, version);
                    case ExcelFillStyle.LightGray:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightGray(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 2], 4, 2, version);
                    case ExcelFillStyle.Gray125:
                        return new PdfTilingPattern(objectNumber, new PdfPatternGray125(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case ExcelFillStyle.Gray0625:
                        return new PdfTilingPattern(objectNumber, new PdfPatternGray0625(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 8, 4], 8, 4, version);
                    case ExcelFillStyle.DarkVertical:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkVertical(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 1], 2, 1, version);
                    case ExcelFillStyle.DarkHorizontal:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkHorizontal(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 1, 2], 1, 2, version);
                    case ExcelFillStyle.DarkDown:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkDown(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 11.3137, 11.3137], 4, 4, version);
                    case ExcelFillStyle.DarkUp:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkUp(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 11.3137, 11.3137], 4, 4, version);
                    case ExcelFillStyle.DarkGrid:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkGrid(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 2, 2], 2, 2, version);
                    case ExcelFillStyle.DarkTrellis:
                        return new PdfTilingPattern(objectNumber, new PdfPatternDarkTrellis(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case ExcelFillStyle.LightVertical:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightVertical(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 1, 0.5d], 1, 0.5, version);
                    case ExcelFillStyle.LightHorizontal:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightHorizontal(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 0.5d, 1], 0.5d, 1, version);
                    case ExcelFillStyle.LightDown:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightDown(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case ExcelFillStyle.LightUp:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightUp(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case ExcelFillStyle.LightGrid:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightGrid(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                    case ExcelFillStyle.LightTrellis:
                        return new PdfTilingPattern(objectNumber, new PdfPatternLightTrellis(CellFillData.PatternColor, CellFillData.BackgroundColor), [0, 0, 4, 4], 4, 4, version);
                }
            }
            return null;
        }
    }
}
