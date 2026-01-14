using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class DrawingBaseItem : DrawingBase
    {
        public DrawingBaseItem(ExcelShapeBase drawing) : base(drawing)
        {
            ImportEpplusDrawing(drawing);
        }

        internal void ImportEpplusDrawing(ExcelShapeBase drawing)
        {
            Bounds.Left = drawing.GetPixelLeft();
            Bounds.Top = drawing.GetPixelTop();
            Bounds.Width = drawing.GetPixelWidth();
            Bounds.Height = drawing.GetPixelHeight();

            SetDrawingPropertiesFill(drawing.Fill, null);
            SetDrawingPropertiesBorder(drawing.Border, null, drawing.Border != null);
        }

        public override RenderItemType Type => RenderItemType.Group;

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left; it = Bounds.Top; ir = Bounds.Right; ib = Bounds.Bottom;
        }
    }
}
