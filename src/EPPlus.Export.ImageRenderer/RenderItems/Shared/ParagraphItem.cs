using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class ParagraphItem : RenderItem
    {
        ITextMeasurerWrap _measurer;

        double _leftMargin;
        double _rightMargin;
        protected eTextAlignment _hAlign;

        eDrawingTextLineSpacing _lsType;
        double _lineSpacingAscendantOnly;
        double? _lsMultiplier = null;
        internal bool IsFirstParagraph { get; private set; }
        List<string> _paragraphLines = new List<string>();
        protected List<string> _textRunDisplayText = new List<string>();

        TextFragmentCollection _textFragments;
        internal protected MeasurementFont _paragraphFont;

        internal double ParagraphLineSpacing { get; private set; }
        internal List<TextRunItem> Runs { get; set; } = new List<TextRunItem>();

        public ParagraphItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {
            Bounds.Name = "Paragraph";
            var defaultFont = new MeasurementFont { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };
            _paragraphFont = defaultFont;
        }

        public ParagraphItem(DrawingBase renderer, BoundingBox parent, ExcelDrawingParagraph p, string textIfEmpty=null) : base(renderer, parent)
        {
            IsFirstParagraph = p == p._paragraphs[0];

            if (p.DefaultRunProperties.Fill != null && p.DefaultRunProperties.Fill.IsEmpty == false)
            {
                if(IsFirstParagraph)
                {
                    SetDrawingPropertiesFill(p.DefaultRunProperties.Fill, null);
                }
                else
                {
                    //Drawingproperties has fallback to firstDefault but excel does not display it so we should not either.
                    if(p.DefaultRunProperties != p._paragraphs.FirstDefaultRunProperties)
                    {
                        SetDrawingPropertiesFill(p.DefaultRunProperties.Fill, null);
                    }
                    else
                    {
                        //Use shape fill somehow
                        //Maybe use a name property for fallback theme accent1 color?
                    }
                }
            }

            //---Initialize Bounds / Margins-- -
            Bounds.Name = "Paragraph";

            var indent = 48 * p.IndentLevel;
            _leftMargin = p.LeftMargin + p.Indent + indent;
            _rightMargin = p.RightMargin;

            _hAlign = p.HorizontalAlignment;

            Bounds.Left = GetAlignmentHorizontal(_hAlign);
            Bounds.Width = parent.Width - p.RightMargin - p.LeftMargin;

            //---Get measurer---
            _measurer = p._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;

            //---Calculate linespacing---
            int numLines = _paragraphLines.Count;
            _lsType = p.LineSpacing.LineSpacingType;
            ParagraphLineSpacing = GetParagraphLineSpacingInPixels(p.LineSpacing.Value, _measurer);

            //---Initialize / calculate lines and runs---
            //measurer must be set before AddLinesAndRichText
            _paragraphFont = p.DefaultRunProperties.GetMeasureFont();
            _measurer.SetFont(_paragraphFont);

            AddLinesAndTextRuns(p, textIfEmpty);
        }

        private double GetParagraphLineSpacingInPixels(double spacingValue, ITextMeasurerWrap fmExact)
        {
            if (_lsType == eDrawingTextLineSpacing.Exactly)
            {
                if (IsFirstParagraph)
                {
                    _lineSpacingAscendantOnly = spacingValue.PointToPixel();
                }
                return spacingValue.PointToPixel();
            }
            else
            {
                var multiplier = (spacingValue / 100);
                _lsMultiplier = multiplier;
                if (IsFirstParagraph)
                {
                    _lineSpacingAscendantOnly = multiplier * fmExact.GetBaseLine().PointToPixel();
                }
                return multiplier * fmExact.GetSingleLineSpacing().PointToPixel();
            }
        }


        /// <summary>
        /// DisplayString is the text altered for display with respect to bounds etc.
        /// Containing line breaks appropriate for the given container
        /// </summary>
        /// <param name="origTxtRun"></param>
        /// <param name="displayText"></param>
        internal protected void AddRenderItemTextRun(ExcelParagraphTextRunBase origTxtRun, string displayText, double startingX)
        {
            //Create object of type
            var targetTxtRun = CreateTextRun(origTxtRun, Bounds, displayText);
            targetTxtRun.LineSpacingPerNewLine = ParagraphLineSpacing;

            if (Runs.Count == 0 && IsFirstParagraph == true)
            {
                targetTxtRun.BaseLineSpacing = _lineSpacingAscendantOnly;
            }

            //If there are multiple sizes/multiple fonts with multiple sizes
            if (_lsMultiplier.HasValue)
            {
                var runFont = origTxtRun.GetMeasurementFont();
                _measurer.SetFont(runFont);
                targetTxtRun.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
                targetTxtRun.BaseLineSpacing = _lsMultiplier.Value * _measurer.GetBaseLine().PointToPixel(true);
                //Reset measurer font
                _measurer.SetFont(_paragraphFont);
            }

            targetTxtRun.Bounds.Left = startingX;

            targetTxtRun.GetBounds(out double l, out double t, out double r, out double b);


            Runs.Add(targetTxtRun);
        }

        public void AddText(string text, ExcelTextFont font)
        {
            var measurer = new FontMeasurerTrueType();
            var displayText = measurer.MeasureAndWrapText(text, font.GetMeasureFont(), Bounds.Parent.Size.X);
            var container = CreateTextRun(text, font, Bounds, string.Join("\r\n", displayText.ToArray()));
            container.BaseLineSpacing = _lineSpacingAscendantOnly;
            container.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
            Runs.Add(container);

            container.Bounds.Name = $"Container{Runs.Count}";
        }

        /// <summary>
        /// Log linebreak positions and sizes of the runs
        /// So that we can easily know what textfragment is on what line and what size it has later
        /// </summary>
        /// <param name="runs"></param>
        void GenerateTextFragments(ExcelDrawingTextRunCollection runs)
        {
            List<string> runContents = new List<string>();
            List<float> fontSizes = new List<float>();
            
            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();

                runContents.Add(txtRun.Text);
                fontSizes.Add(runFont.Size);
            }

            _textFragments = new TextFragmentCollection(runContents, fontSizes);
        }

        List<string> GetWrappedText(ExcelDrawingTextRunCollection runs, TextFragmentCollection fragments)
        {
            var ttMeasurer = (FontMeasurerTrueType)_measurer;
            List<MeasurementFont> fonts = new List<MeasurementFont>();

            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();
                fonts.Add(runFont);
            }

            var maxSizePoints = Math.Round(Bounds.Width, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            return ttMeasurer.WrapMultipleTextFragments(fragments, fonts, maxSizePoints);
        }

        private void AddLinesAndTextRuns(ExcelDrawingParagraph p, string textIfEmpty)
        {
            //Log line positions and run sizes
            GenerateTextFragments(p.TextRuns);
            //Calculate line breaks and/or wrapping to know how the text should be displayed
            CalculateDisplayText(p, _textFragments);

            string currentLine = _paragraphLines[0];
            double widthOfCurrentLine = 0;
            double largestFontSizeCurrentLine = 0;
            int idxLargestFontSize = 0;
            if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
            {
                AddText(textIfEmpty, p.DefaultRunProperties);
            }
            else
            {
                for (int i = 0; i < p.TextRuns.Count; i++)
                {
                    if (p.TextRuns[i].FontSize > largestFontSizeCurrentLine)
                    {
                        largestFontSizeCurrentLine = p.TextRuns[i].FontSize;
                        idxLargestFontSize = i;
                    }

                    AddRenderItemTextRun(p.TextRuns[i], _textRunDisplayText[i], widthOfCurrentLine);
                    var lastAdded = Runs.Last();

                    //We are on a new line
                    if (lastAdded.YIncreasePerLine.Count > 1)
                    {
                        Runs[idxLargestFontSize].GetBounds(out double l, out double t, out double r, out double b);
                        Bounds.Height += Runs[idxLargestFontSize].Bounds.Height;
                        widthOfCurrentLine = Runs.Last().PerLineWidth.Last();

                        idxLargestFontSize = i;
                        largestFontSizeCurrentLine = p.TextRuns[i].FontSize;
                    }
                    else
                    {
                        widthOfCurrentLine += Runs.Last().PerLineWidth.Last();
                        if (i == p.TextRuns.Count - 1)
                        {
                            Runs[idxLargestFontSize].GetBounds(out double l, out double t, out double r, out double b);
                            Bounds.Height += Runs[idxLargestFontSize].Bounds.Height;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Use textfragments to calculate wrapping/line-breaks
        /// </summary>
        /// <param name="p"></param>
        private void CalculateDisplayText(ExcelDrawingParagraph p, TextFragmentCollection fragments)
        {
            //Gets the individual lines free of any line breaks
            if (p._paragraphs.WrapText != eTextWrappingType.None)
            {
                _paragraphLines = GetWrappedText(p.TextRuns, fragments);
            }
            else
            {
                //Using Regex to avoid empty lines in windows.
                //Which the alternative "paragraph.Text.Split(new [] '\r' '\n')" would result in.
                _paragraphLines = Regex.Split(p.Text, "\r\n|\r|\n").ToList();
            }
            //Gets each actual text run, with linebreak symbols
            _textRunDisplayText = fragments.GetFragmentsWithFinalLineBreaks();
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

        //internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        //{
        //    il = Bounds.Left + _leftMargin;
        //    it = Bounds.Top;
        //    ir = Bounds.Right - _rightMargin;
        //    ib = Bounds.Bottom;
        //}


        /// <summary>
        /// Type of textrun defined by child type
        /// </summary>
        /// <param name="run"></param>
        /// <param name="parent"></param>
        /// <param name="DisplayString"></param>
        /// <returns></returns>
        internal abstract TextRunItem CreateTextRun(ExcelParagraphTextRunBase run, BoundingBox parent, string displayText);
        internal abstract TextRunItem CreateTextRun(string text, ExcelTextFont font, BoundingBox parent, string displayText);
    }
}
