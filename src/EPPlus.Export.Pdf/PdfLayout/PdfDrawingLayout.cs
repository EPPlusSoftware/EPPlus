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
using OfficeOpenXml.Drawing;
using EPPlus.Graphics;


namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfDrawingLayout : Transform
    {
        public ExcelDrawing Drawing;
        public PdfDrawingLayout(ExcelDrawing drawing, double x, double y, double width, double height)
            : base(x,y,width,height)
        {
            Drawing = drawing;
        }
    }
}
