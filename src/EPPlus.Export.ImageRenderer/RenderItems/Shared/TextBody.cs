using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    /// <summary>
    /// Margin left = X
    /// Margin right = Y
    /// </summary>
    internal class TextBody : RenderItem
    {
        /// <summary>
        /// Shorthand for Bounds.Width
        /// </summary>
        internal double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }

        /// <summary>
        /// Shorthand for Bounds.Height
        /// </summary>
        internal double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }

        internal string FontColorString { get; set; }

        internal eTextAnchoringType VerticalAlignment = eTextAnchoringType.Top;

        internal List<ParagraphContainer> Paragraphs = new List<ParagraphContainer>();

        public bool AllowOverflow;

        internal bool WrapText = true;

        private FontMeasurerTrueType _measurer = null;

        private double paragraphStartPosY = 0;

        public TextBody(BoundingBox parent) : base()
        {
            Bounds.Parent = parent;

            Bounds.Width = parent.Width;
            Bounds.Height = parent.Height;
        }

        public void AddParagraph(string text, FontMeasurerTrueType measurer)
        {
            if(_measurer == null && measurer != null)
            {
                _measurer = measurer;
            }

            if (_measurer != null)
            {
                var paragraph = new ParagraphContainer(Bounds);

                Paragraphs.Add(paragraph);

                paragraph.AddText(text, measurer);

                paragraph.Bounds.transform.Name = $"Container{Paragraphs.Count}";
                return;
            }
            throw new NullReferenceException($"The FontMeasurer: {_measurer} object is null. Use SetMeasurer before adding text");
        }

        public void ImportParagraph(ExcelDrawingParagraph item)
        {
            var posY = GetAlignmentVertical();

            SetDrawingPropertiesFill(item.DefaultRunProperties.Fill, null);

            var measureFont = item.DefaultRunProperties.GetMeasureFont();
            bool isFirst = Paragraphs.Count == 0;

            var paragraph = new ParagraphContainer(item, Bounds);
        }

        public void ImportTextBody(ExcelTextBody body)
        {
            double l, r, t, b;
            body.GetInsetsOrDefaults(out l, out t, out r, out b);

            Bounds.Left = l.PointToPixel();
            Bounds.Top = t.PointToPixel();
            Bounds.Right = r.PointToPixel();
            Bounds.Bottom = b.PointToPixel();

            foreach (var paragraph in body.Paragraphs)
            {
                ImportParagraph(paragraph);
            }
        }

        public string GetContent()
        {
            return Paragraphs[0].GetContent();
            //string.Join(Paragraphs[}])
        }

        public void AddText(string text, FontMeasurerTrueType measurer)
        {
            if (Paragraphs.Count <= 0)
            {
                AddParagraph(text, measurer);
            }
            else
            {
                Paragraphs.Last().AddText(text, measurer);
            }
        }

        public void SetMeasurer(FontMeasurerTrueType fontMeasurer)
        {
            _measurer = fontMeasurer;
        }

        /// <summary>
        /// Same as X or Left
        /// </summary>
        internal double LeftMargin { get { return Bounds.Left; } set { Bounds.Left = value; } }

        /// <summary>
        /// Same as Y or Top
        /// </summary>
        internal double TopMargin { get { return Bounds.Y; } set { Bounds.Y = value; } }

        double _rightMargin = 0;

        /// <summary>
        /// Setting this property also changes the width of the textbody
        /// </summary>
        internal double RightMargin { get { return _rightMargin; } set { _rightMargin = value; Bounds.Width = Bounds.Parent.Width - value; } }

        double _bottomMargin = 0;
        /// <summary>
        /// Setting this property also changes the height of the textbody
        /// </summary>
        internal double BottomMargin { get { return _bottomMargin; } set { _bottomMargin = value; Bounds.Height = Bounds.Parent.Height - value; } }

        public override RenderItemType Type => throw new NotImplementedException();

        double? _alignmentY = null;

        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <param name="fontSizeInPixels"></param>
        /// <returns></returns>
        private double GetAlignmentVertical()
        {
            double alignmentY = 0;

            switch (VerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    alignmentY = Bounds.Y;
                    break;
                case eTextAnchoringType.Center:
                    var adjustedHeight =  Bounds.Y + (Bounds.Height / 2);

                    alignmentY = adjustedHeight;
                    break;
                case eTextAnchoringType.Bottom:
                    alignmentY = Bounds.Height;
                    break;
            }

            _alignmentY = alignmentY;

            return _alignmentY.Value;
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left; it = Bounds.Top; ir = Bounds.Right; ib = Bounds.Bottom;
        }

        public override void Render(StringBuilder sb)
        {
            //No rendering in generic position and styling class
        }
    }
}
