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
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAxisTextBoxes : ChartDrawingObject
    {
        internal override Color? DefaultFillColor { get; }

        internal ChartAxisTextBoxes(ChartRenderer chart) : base(chart)
        {
            DefaultFillColor = Color.Transparent;
        }

        internal List<DrawingTextBox> TextBoxes
        {
            get;
            set;
        }=new List<DrawingTextBox>();

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            if (TextBoxes != null && TextBoxes.Count > 0)
            {
                foreach (var tb in TextBoxes)
                {
                    tb.AppendRenderItems(renderItems);
                }
            }

        }
    }
}