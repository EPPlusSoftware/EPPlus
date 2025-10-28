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
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer
{
    internal abstract class DrawingBase
    {
         internal DrawingBase(ExcelDrawing drawing)
        {
            drawing.GetSizeInPixels(out int width, out int height);
            Drawing = drawing;
            Size = new DrawingSize(width, height);
            TextMeasurer = drawing._drawings._package.Settings.TextSettings.PrimaryTextMeasurer;
        }
        public ExcelDrawing Drawing { get; }
        internal ITextMeasurer TextMeasurer { get; }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        public DrawingSize Size { get; internal set; }
        public abstract void Render(StringBuilder sb);
    }
}