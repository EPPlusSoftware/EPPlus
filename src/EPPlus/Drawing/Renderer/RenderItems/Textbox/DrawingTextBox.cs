using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using OfficeOpenXml.Style;
using System;

namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    public class DrawingTextBox : RenderTextbox
    {
        ExcelDrawing _drawing;

        /// <summary>
        /// Creates a text box at a given position.
        /// </summary>
        /// <param name="renderContext">The render context of the renderer creating the text box. Decides the font render target.</param>
        internal DrawingTextBox(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN)
            : base(parent, left, top, width, height, maxWidth, maxHeight)
        {
            Init(renderContext, drawing, parent, maxWidth, maxHeight);
            Left = left;
            Top = top;
        }

        /// <summary>
        /// Creates a text box to be positioned later.
        /// </summary>
        /// <param name="renderContext">The render context of the renderer creating the text box. Decides the font render target.</param>
        internal DrawingTextBox(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight)
            : base(parent, maxWidth, maxHeight)
        {
            Init(renderContext, drawing, parent, maxWidth, maxHeight);
        }

        private void Init(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight)
        {
            if (renderContext == null)
                throw new ArgumentNullException("renderContext");

            Parent = parent;
            _drawing = drawing;
            //The context must come from the renderer, not the workbook, so the render target is preserved.
            TextBody = new DrawingTextBody(renderContext, drawing, _marginGroup, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;
        }

        internal void AddText(string text = null)
        {
            TextBody.AddParagraph(text);
        }

        DrawingTextBody _textBody;

        public DrawingTextBody GetTextBody()
        {
            return (DrawingTextBody)TextBody;
        }

        public void SetDrawingTextBody(DrawingTextBody tb)
        {
            TextBody = tb;
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

            _textBody.ImportTextBodyAndParagraphs(body, horizontalDefault);
        }

        internal void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            _textBody.ImportParagraph(item, startingY, text);
        }
    }
}