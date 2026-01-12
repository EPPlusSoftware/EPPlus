using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class ParagraphContainer : RenderItem
    {
        ITextMeasurerWrap _measurer;

        double _leftMargin;
        double _rightMargin;
        protected eTextAlignment _hAlign;

        eDrawingTextLineSpacing _lsType;
        internal double ParagraphLineSpacing { get; private set; }
        double _lineSpacingAscendantOnly;
        double? _lsMultiplier = null;
        bool _isFirstParagraph;

        public override RenderItemType Type => RenderItemType.Text;

        List<string> _paragraphLines = new List<string>();
        protected List<string> _textRunContent = new List<string>();

        internal List<FontWrapContainer> Runs = new List<FontWrapContainer>();
        TextFragmentCollection _textFragments;
        internal List<TextRunRenderItem> _textRunItems;

        internal protected MeasurementFont _paragraphFont;

        public ParagraphContainer() : base()
        {

        }

        public ParagraphContainer(BoundingBox parent)
        {
            Bounds.transform.Name = "Paragraph";
            Bounds.Parent = parent;
        }

        public ParagraphContainer(ExcelDrawingParagraph p, BoundingBox parent) : base()
        {
            //---Initialize Bounds/Margins---
            Bounds.transform.Name = "Paragraph";
            Bounds.Parent = parent;

            var indent = 48 * p.IndentLevel;
            _leftMargin = p.LeftMargin + p.Indent + indent;
            _rightMargin = p.RightMargin;

            _hAlign = p.HorizontalAlignment;

            Bounds.X = GetAlignmentHorizontal(_hAlign);
            Bounds.Width = parent.Width - p.RightMargin;

            //---Initialize / calculate lines and runs---
            //measurer be set before InitializeLinesAndRichText

            _measurer = p._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
            _paragraphFont = p.DefaultRunProperties.GetMeasureFont();
            _measurer.SetFont(_paragraphFont);

            InitializeLinesAndRichText(p);

            _isFirstParagraph = p == p._paragraphs[0];
            
            //---Calculate linespacing---
            int numLines = _paragraphLines.Count;
            _lsType = p.LineSpacing.LineSpacingType;
            ParagraphLineSpacing = GetParagraphLineSpacingInPixels(p.LineSpacing.Value, _measurer);

            _textRunItems = new List<TextRunRenderItem>();
        }

        private double GetParagraphLineSpacingInPixels(double spacingValue, ITextMeasurerWrap fmExact)
        {
            if (_lsType == eDrawingTextLineSpacing.Exactly)
            {
                if (_isFirstParagraph)
                {
                    _lineSpacingAscendantOnly = spacingValue.PointToPixel();
                }
                return spacingValue.PointToPixel();
            }
            else
            {
                var multiplier = (spacingValue / 100);
                _lsMultiplier = multiplier;
                if (_isFirstParagraph)
                {
                    _lineSpacingAscendantOnly = multiplier * fmExact.GetBaseLine().PointToPixel();
                }
                return multiplier * fmExact.GetSingleLineSpacing().PointToPixel();
            }
        }


        /// <summary>
        /// Each specific container must specify textrun type within this
        /// </summary>
        /// <param name="origTxtRun"></param>
        /// <param name="runDisplayString"></param>
        internal protected abstract void AddTextRun(ExcelParagraphTextRunBase origTxtRun, string runDisplayString);

        /// <summary>
        /// DisplayString is the text altered for display with respect to bounds etc.
        /// Containing line breaks appropriate for the given container
        /// </summary>
        /// <param name="origTxtRun"></param>
        /// <param name="runDisplayString"></param>
        internal protected void AddTextRun<T>(ExcelParagraphTextRunBase origTxtRun, string runDisplayString) where T : TextRunRenderItem
        {
            var maxWidth = Bounds.Width;

            //Specify type
            var textRunType = typeof(T);

            //Create object of type
            var targetTxtRun = (T)Activator.CreateInstance(textRunType, origTxtRun, Bounds, runDisplayString);

            if (_textRunItems.Count == 0 && _isFirstParagraph == true)
            {
                targetTxtRun.BaseLineSpacing = _lineSpacingAscendantOnly;
                targetTxtRun.LineSpacingPerNewLine = ParagraphLineSpacing;
            }
            else
            {
                targetTxtRun.LineSpacingPerNewLine = ParagraphLineSpacing;

                //If there are multiple sizes/multiple fonts with multiple sizes
                if (_lsMultiplier.HasValue)
                {
                    var runFont = origTxtRun.GetMeasurementFont();
                    _measurer.SetFont(runFont);
                    targetTxtRun.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
                    //Reset measurer font
                    _measurer.SetFont(_paragraphFont);
                }
            }

            //No need for most of the actual output variables anymore
            targetTxtRun.GetBounds(out double l, out double t, out double r, out double b);

            Bounds.Bottom += b;

            _textRunItems.Add(targetTxtRun);
        }

        public void AddText(string text, FontMeasurerTrueType measurer)
        {
            var container = new FontWrapContainer(measurer);
            container.Parent = Bounds;

            Runs.Add(container);

            container.transform.Name = $"Container{Runs.Count}";

            container.SetContent(text);
        }

        public string GetContent()
        {
            StringBuilder sb = new StringBuilder();

            foreach (var item in Runs)
            {
                sb.Append(item.GetContent());
            }

            return sb.ToString();
        }

        List<string> GetWrappedText(ExcelDrawingTextRunCollection runs)
        {
            var ttMeasurer = (FontMeasurerTrueType)_measurer;

            List<string> runContents = new List<string>();
            List<MeasurementFont> fonts = new List<MeasurementFont>();

            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasureFont();

                runContents.Add(txtRun.Text);
                fonts.Add(runFont);
            }

            _textFragments = new TextFragmentCollection(runContents);
            var pointWidth = Bounds.Width.PixelToPoint();
            return ttMeasurer.WrapMultipleTextFragments(_textFragments, fonts, pointWidth);
        }

        private void InitializeLinesAndRichText(ExcelDrawingParagraph paragraph)
        {
            if (paragraph._paragraphs.WrapText != eTextWrappingType.None)
            {
                //Gets the individual lines free of any line breaks
                _paragraphLines = GetWrappedText(paragraph.TextRuns);

                //Gets each actual text run, including linebreak symbols
                _textRunContent = _textFragments.GetFragmentsWithFinalLineBreaks();
            }
            else
            {
                //Gets the individual lines free of any line breaks

                //Using Regex to avoid empty lines in windows.
                //Which the alternative "paragraph.Text.Split(new [] '\r' '\n')" would result in.
                _paragraphLines = Regex.Split(paragraph.Text, "\r\n|\r|\n").ToList();

                //Gets each actual text run, including linebreak symbols
                foreach (var run in paragraph.TextRuns)
                {
                    _textRunContent.Add(run.Text);
                }
            }
        }

        internal double GetAlignmentHorizontal(eTextAlignment txAlignment)
        {
            var area = Bounds;
            double x = 0;
            switch (txAlignment)
            {
                case eTextAlignment.Left:
                default:
                    x = area.Left + _leftMargin;
                    break;
                case eTextAlignment.Center:
                    x = (area.Right / 2) + _leftMargin - _rightMargin;
                    break;
                case eTextAlignment.Right:
                    x = area.Right - _rightMargin;
                    break;
            }

            return TextUtils.RoundToWhole(x);
        }

        //public override void Render(StringBuilder sb)
        //{
        //    throw new NotImplementedException();
        //}

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left + _leftMargin;
            it = Bounds.Top;
            ir = Bounds.Right - _rightMargin;
            ib = Bounds.Bottom;
        }
    }
}
