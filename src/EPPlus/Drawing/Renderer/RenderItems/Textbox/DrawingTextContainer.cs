using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.Textbox;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Renderer.RenderItems.Textbox
{
    public class DrawingTextContainer : TextContainer
    {
        //internal DrawingTextContainer(BoundingBox parent) : base(parent)
        //{

        //}

        /// <summary>
        /// Creates a text box at a given position.
        /// </summary>
        /// <param name="renderContext">The render context of the renderer creating the text box. Decides the font render target.</param>
        internal DrawingTextContainer
          (
            RenderContext renderContext, 
            ExcelDrawing drawing, 
            BoundingBox parent, 
            MeasurementFont defaultFont, 
            double left, double top, double width, double height,
            double maxWidth = double.NaN, double maxHeight = double.NaN
          ):base(parent, renderContext)
        {
            Init(renderContext, drawing, parent, maxWidth, maxHeight, defaultFont);
            Left = left;
            Top = top;
        }

        internal DrawingTextContainer(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, MeasurementFont defaultFont, double maxWidth, double maxHeight)
            : base(parent, renderContext)
        {
            Init(renderContext, drawing, parent, maxWidth, maxHeight, defaultFont);
        }

        private void Init(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight, MeasurementFont defaultFont)
        {
            if (renderContext == null)
                throw new ArgumentNullException("renderContext");
            //The context must come from the renderer, not the workbook, so the render target is preserved.
            TextBody = new DrawingTextBody(renderContext, drawing, _marginGroup, defaultFont, true)
            {
                MaxWidth = maxWidth,
                MaxHeight = maxHeight
            };

            Parent = parent;
        }

        public override RenderTextBody TextBody { get { return _textBody; } set { _textBody = (DrawingTextBody)value; } }

        internal void ImportTextBodyAndParagraphs(ExcelTextBody body, bool useDefaults = true, ExcelHorizontalAlignment horizontalDefault = ExcelHorizontalAlignment.Left)
        {
            double l, r, t, b;
            if (useDefaults)
            {
                body.GetInsetsOrDefaults(out l, out t, out r, out b);
            }
            else
            {
                body.GetInsetsInPoints(out l, out t, out r, out b);
            }
            LeftMargin = l;
            TopMargin = t;
            RightMargin = r;
            BottomMargin = b;

            ((DrawingTextBody)_textBody).ImportTextBodyAndParagraphs(body, horizontalDefault);
            CalculateMargin();
        }

        internal void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            ((DrawingTextBody)_textBody).ImportParagraph(item, startingY, text);
            CalculateMargin();
        }
    }
}
