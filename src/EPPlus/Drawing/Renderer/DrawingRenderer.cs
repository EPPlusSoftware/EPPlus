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

using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Export.Utils;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Renderer
{
    internal abstract class DrawingRenderer
    {
        internal DrawingRenderer(ExcelDrawing drawing)
        {
            Drawing = drawing;
            Bounds = drawing.GetBoundingBox();

            var wb = drawing._drawings.Worksheet.Workbook;
            Theme = wb.ThemeManager.GetOrCreateTheme();

            var shaper = OpenTypeFonts.GetTextShaper(Theme.FontScheme.MajorFont[0].Typeface);
            TextMeasurer = new OpenTypeFontTextMeasurer(shaper);
        }


        internal DrawingRenderer()
        {
            //Drawing = drawing;
            //Bounds = drawing.GetBoundingBox();

            //var wb = drawing._drawings.Worksheet.Workbook;
            //Theme = wb.ThemeManager.GetOrCreateTheme();
        }

        internal readonly StyleCache _styleCache = new StyleCache();
        public ExcelDrawing Drawing { get; }
        public ExcelTheme Theme { get;}
        public ExcelWorkbook Workbook => Drawing._drawings.Worksheet.Workbook;
        internal ITextMeasurer TextMeasurer { get; }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        internal BoundingBox Bounds = new BoundingBox();
    }
}