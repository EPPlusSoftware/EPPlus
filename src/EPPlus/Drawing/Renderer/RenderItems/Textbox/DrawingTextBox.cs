using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    public class DrawingTextBox : RenderTextbox
    {
        ExcelDrawing _drawing;
        internal DrawingTextBox(ExcelDrawing drawing, BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN) : base(parent, left, top, width, height, maxWidth, maxHeight)
        {
            Init(drawing, parent, maxWidth, maxHeight);
            Left = left;
            Top = top;
        }

        private void Init(ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight) 
        {
            Parent = parent;
            _drawing= drawing;
            TextBody = new DrawingTextBody(drawing, _marginGroup.Bounds, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;
        }

        internal DrawingTextBox(ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight) : base(parent, maxWidth, maxHeight)
        {
            Init(drawing, parent, maxWidth, maxHeight);
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

        //internal void AddText(double startingY, string text = null)
        //{
        //    TextBody.AddParagraph(startingY, text);
        //}
    }
}
