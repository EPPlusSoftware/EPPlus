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
        //protected ExcelWorkbook _wb;
        protected ExcelDrawing _drawing;
        protected ExcelTheme _theme;

        internal DrawingBase(ExcelDrawing drawing)
        {
            //drawing.GetSizeInPixels(out int width, out int height);
            Drawing = drawing;
            //Bounds = new DrawingSize(width, height);
            Bounds = drawing.GetBoundingBox();
            //TextMeasurer = drawing._drawings._package.Settings.TextSettings.PrimaryTextMeasurer;

            var wb = drawing._drawings.Worksheet.Workbook;
            _theme = wb.ThemeManager.GetOrCreateTheme();
        }
        public ExcelDrawing Drawing { get; }
        //internal ITextMeasurer TextMeasurer { get; }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        //public DrawingSize Bounds { get; internal set; }
        internal BoundingBox Bounds = new BoundingBox();
    }
}