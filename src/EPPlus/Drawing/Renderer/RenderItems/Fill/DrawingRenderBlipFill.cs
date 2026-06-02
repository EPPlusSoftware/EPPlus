using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.RenderItems.Fill
{
    public class DrawingRenderBlipFill : RenderBlipFill
    {
        //private ExcelDrawingBlipFill _blipFill;

        internal DrawingRenderBlipFill(ExcelDrawingBlipFill blipFill)
        {
            //_blipFill = blipFill;
            ImageBytes = blipFill.Image.ImageBytes;
            ImageBounds = new BoundingBox(blipFill.Image.Bounds.Width, blipFill.Image.Bounds.Height);
            ContentType = blipFill.ContentType;
            Stretch = blipFill.Stretch;
            StretchOffset = blipFill.StretchOffset.AsOffsetRectangle();
            Tile = blipFill.Tile.AsFillTile();
            
        }
    }
}
