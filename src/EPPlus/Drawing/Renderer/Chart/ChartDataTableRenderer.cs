/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  20/08/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlus.DrawingRenderer.RenderItems;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class ChartDataTableRenderer : ChartDrawingObject
    {
        internal ChartDataTableRenderer(ChartRenderer svgChart) : base(svgChart)
        {
             var chartDataTable = svgChart.Chart.PlotArea.DataTable;
            
        }
        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            base.AppendRenderItems(renderItems);
        }
        
    }
}