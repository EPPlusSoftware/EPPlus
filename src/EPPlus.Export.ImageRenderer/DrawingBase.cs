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

using EPPlus.Export.ImageRenderer;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer
{
    internal abstract class DrawingBase
    {
        internal DrawingBase(ExcelDrawing drawing)
        {
            Drawing = drawing;
            Bounds = drawing.GetBoundingBox();

            var wb = drawing._drawings.Worksheet.Workbook;
            Theme = wb.ThemeManager.GetOrCreateTheme();
        }


        internal DrawingBase()
        {
            //Drawing = drawing;
            //Bounds = drawing.GetBoundingBox();

            //var wb = drawing._drawings.Worksheet.Workbook;
            //Theme = wb.ThemeManager.GetOrCreateTheme();
        }

        public ExcelDrawing Drawing { get; }
        public ExcelTheme Theme { get;}
        public ExcelWorkbook Workbook => Drawing._drawings.Worksheet.Workbook;
        internal ITextMeasurer TextMeasurer { get; }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        internal BoundingBox Bounds = new BoundingBox();
    }
}