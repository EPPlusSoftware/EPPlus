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
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.PdfExport.Data.Dictionaries
{
    internal class PdfShadingResource : PdfResource
    {
        internal int objectNumber;
        internal PdfCellFillData CellFillData;

        public PdfShadingResource(int labelNumber, PdfCellFillData cellFillData)
            : base("Sh", labelNumber)
        {
            CellFillData = cellFillData;
        }

        public PdfShading GetShadingObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
            if (CellFillData.GradientFillData != null)
            {
                if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Linear)
                {
                    var pas = new PdfAxialShading(objectNumber, CellFillData.GradientFillData, version);
                    pas.Coords = CellFillData.GradientFillData.coords;
                    return pas;
                }
                else if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Path)
                {
                    var prs = new PdfRadialShading(objectNumber, CellFillData.GradientFillData, version);
                    prs.Coords = CellFillData.GradientFillData.coords;
                    return prs;
                }
            }
            return null;
        }

        public PdfShadingPattern GetShadingPatternObject(int patternObjectNumber, int shadingObjectNumber, int version = 0)
        {
            objectNumber = shadingObjectNumber;
            var shadingPattern = new PdfShadingPattern(patternObjectNumber, shadingObjectNumber);
            shadingPattern.Matrix = CellFillData.GradientFillData.matrix;
            return shadingPattern;
        }

    }
}
