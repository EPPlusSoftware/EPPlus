using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using EPPlusColorConverter = OfficeOpenXml.Utils.TypeConversion.ColorConverter;

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
        internal TextBodyItem ParentTextBody { get; set; }
        internal double ParagraphLineSpacing { get; private set; }
        internal List<TextRunItem> Runs { get; set; } = new List<TextRunItem>();

        public ParagraphItem(TextBodyItem textBody, DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {
            ParentTextBody = textBody;
            Bounds.Name = "Paragraph";
            var defaultFont = new MeasurementFont { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };
            _paragraphFont = defaultFont;
        }

        public ParagraphItem(TextBodyItem textBody, DrawingBase renderer, BoundingBox parent, ExcelDrawingParagraph p, string textIfEmpty=null) : base(renderer, parent)
        {
            ParentTextBody = textBody; 
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
                        var fc = EPPlusColorConverter.GetThemeColor(DrawingRenderer.Theme.ColorScheme.Light1);
                        fc = ColorUtils.GetAdjustedColor(PathFillMode.Norm, fc);
                        FillColor = "#" + fc.ToArgb().ToString("x8").Substring(2);
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

        internal protected TextRunItem AddRenderItemTextRun(ExcelParagraphTextRunBase origTxtRun, string displayText, double startingX, double lineSpacing)
        {
            var targetTxtRun = CreateTextRun(origTxtRun, Bounds, displayText);
            targetTxtRun.lineSpacing = lineSpacing;
            targetTxtRun.Bounds.Left = startingX;

            Runs.Add(targetTxtRun);
            return targetTxtRun;
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
            targetTxtRun.lineSpacing = ParagraphLineSpacing;
            targetTxtRun.Bounds.Left = startingX;
            //if (Runs.Count == 0)
            //{
            //    targetTxtRun.BaseLineSpacing = _lineSpacingAscendantOnly;
            //}

            ////If there are multiple sizes/multiple fonts with multiple sizes
            //if (_lsMultiplier.HasValue)
            //{
            //    var runFont = origTxtRun.GetMeasurementFont();
            //    _measurer.SetFont(runFont);
            //    targetTxtRun.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
            //    targetTxtRun.BaseLineSpacing = _lsMultiplier.Value * _measurer.GetBaseLine().PointToPixel(true);
            //    //Reset measurer font
            //    _measurer.SetFont(_paragraphFont);
            //}

            //targetTxtRun.Bounds.Left = startingX;
            //targetTxtRun.GetBounds(out double l, out double t, out double r, out double b);

            //for (int i = 1; i < targetTxtRun.Lines.Count; i++)
            //{

            //}
            //targetTxtRun.SetPerLineWidths(_textFragments.GetFragmentWidths(fragIdx));

            //lineIdxAfter = currentLineIdx + targetTxtRun.Lines.Count - 1;

            Runs.Add(targetTxtRun);
        }

        public void AddText(string text, ExcelTextFont font)
        {
            var measurer = new FontMeasurerTrueType();
            var displayText = measurer.MeasureAndWrapText(text, font.GetMeasureFont(), ParentTextBody.MaxWidth);
            var container = CreateTextRun(text, font, Bounds, string.Join("\r\n", displayText.ToArray()));
            //container.BaseLineSpacing = _lineSpacingAscendantOnly;
            //container.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
            Runs.Add(container);
            Bounds.Width = container.Bounds.Width + 0.001; //TODO: fix for equal width issue
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

        List<TextLineSimple> GetWrappedTextLines(ExcelDrawingTextRunCollection runs, TextFragmentCollection fragments)
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
            return ttMeasurer.WrapMultipleTextFragmentsToTextLines(fragments, fonts, maxSizePoints);
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
        public List<ParagraphLine> Lines { get; set; }
        private void AddLinesAndTextRuns2(ExcelDrawingParagraph p, string textIfEmpty)
        {
            //var line = new ParagraphLine();
            foreach (var r in p.TextRuns)
            {
                foreach(var text in r.SplitIntoLines())
                {
                    var line = new ParagraphLine();

                    //var textWidth = GetWidth(text, r);
                    //if(line.Width+textWidth > maxWidth)
                    //{

                    //}
                }
            }
        }
        private void AddLinesAndTextRuns(ExcelDrawingParagraph p, string textIfEmpty)
        {
            //Log line positions and run sizes
            GenerateTextFragments(p.TextRuns);
            ////Calculate line breaks and/or wrapping to know how the text should be displayed
            //CalculateDisplayText(p, _textFragments);

            var lines = WrapToSimpleTextLines(p, _textFragments);
            //In points
            double lastDescent = 0;
            bool lineSpacingIsExact = _lsMultiplier.HasValue == false;
            double runLineSpacing = 0;
            double greatestWidth = 0;

            foreach (var line in lines)
            {
                double prevWidth = 0;
                
                if(lineSpacingIsExact == false)
                {
                    runLineSpacing += line.LargestAscent + lastDescent;
                }
                else
                {
                    runLineSpacing += ParagraphLineSpacing;
                }
                if(line.Width > greatestWidth)
                {
                    greatestWidth = line.Width;
                }

                foreach (var rtFragment in line.RtFragments)
                {
                    var displayText = line.GetFragmentText(rtFragment);
                    var runItem = AddRenderItemTextRun(p.TextRuns[rtFragment.Fragidx], displayText, prevWidth, runLineSpacing);
                    runItem.Bounds.Width = rtFragment.Width;
                    prevWidth += rtFragment.Width;
                }

                lastDescent = line.LargestDescent;
            }
            Bounds.Height = runLineSpacing + lastDescent;
            //Bounds.Width = greatestWidth;
            //string currentLine = _paragraphLines[0];
            //int currentLineIdx = 0;

            //double widthOfCurrentLine = 0;
            //double largestFontSizeCurrentLine = 0;
            //int idxLargestFontSize = 0;
            //int firstRunInLineIdx = 0;

            ////var lineSizes = _textFragments.GetLargestFontSizesOfEachLine();
            ////var currentLineSize = lineSizes[currentLineIdx];
            //double lineSpacing = 0;
            //if(_lsMultiplier.HasValue == false)
            //{
            //    //linespacing is exact
            //    lineSpacing = ParagraphLineSpacing;
            //}

            if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
            {
                AddText(textIfEmpty, p.DefaultRunProperties);
                Bounds.Width = Runs.Sum(x=>x.Bounds.Width);
            }
            else
            {
                //for (int i = 0; i < p.TextRuns.Count; i++)
                //{
                //    AddRenderItemTextRun(p.TextRuns[i], _textRunDisplayText[i], widthOfCurrentLine);
                //    var lastAdded = Runs.Last();

                //    lastAdded.SetPerLineWidths(_textFragments.GetFragmentWidths(i));

                //    if (lastAdded.Lines.Count > 1)
                //    {
                //        //Just in case. Should always be empty here
                //        lastAdded.YIncreasePerLine.Clear();

                //        //We are on a new line
                //        for (int j = 0; j < lastAdded.Lines.Count; j++)
                //        {
                //            currentLine = _paragraphLines[currentLineIdx];

                //            if (_lsMultiplier.HasValue)
                //            {
                //                //Add ascent to descent (Add ascent to Nothing for the first run)
                //                lineSpacing += _textFragments.GetAscent(currentLineIdx).PointToPixel() * _lsMultiplier.Value;
                //            }

                //            lastAdded.AddLineSpacing(lineSpacing);

                //            if (_lsMultiplier.HasValue)
                //            {
                //                //Set linespacing to descent
                //                lineSpacing = _textFragments.GetDescent(currentLineIdx).PointToPixel();
                //            }

                //            //Last line in added lines we will continue on if there are more textruns
                //            //Therefore the index will be added to after the next line-break or at the end
                //            if (j < lastAdded.Lines.Count - 1)
                //            {
                //                Bounds.Height += lastAdded.Bounds.Height;
                //                currentLineIdx++;
                //            }
                //        }

                //        widthOfCurrentLine = Runs.Last().PerLineWidth.Last();
                //    }
                //    else
                //    {
                //        widthOfCurrentLine += Runs.Last().PerLineWidth.Last();

                //        //If we are on the last run
                //        if (i == p.TextRuns.Count - 1)
                //        {
                //            if (_lsMultiplier.HasValue)
                //            {
                //                //currentLineIdx++;
                //                //Add ascent to descent (Add ascent to Nothing for the first run)
                //                lineSpacing += _textFragments.GetAscent(currentLineIdx).PointToPixel() * _lsMultiplier.Value;
                //            }
                //            lastAdded.AddLineSpacing(lineSpacing);
                //            Bounds.Height += lastAdded.Bounds.Height;
                //            //Runs[idxLargestFontSize].GetBounds(out double l, out double t, out double r, out double b);
                //            //Bounds.Height += Runs[idxLargestFontSize].Bounds.Height;
                //        }
                //    }
                //    Bounds.Width = widthOfCurrentLine;
                //}
            }
        }

        private void CalculateTextLines(ExcelDrawingParagraph p, TextFragmentCollection fragments)
        {
            WrapToSimpleTextLines(p, fragments);
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
            //var test = fragments.GetFragmentsWithoutLineBreaks();
            //var 
            //var lineToRunMapping = 
        }

        List<TextLineSimple> WrapToSimpleTextLines(ExcelDrawingParagraph p, TextFragmentCollection fragments)
        {
            var ttMeasurer = (FontMeasurerTrueType)_measurer;
            List<MeasurementFont> fonts = new List<MeasurementFont>();

            for (int i = 0; i < p.TextRuns.Count(); i++)
            {
                var txtRun = p.TextRuns[i];
                var runFont = txtRun.GetMeasurementFont();
                fonts.Add(runFont);
            }

            var maxWidthPoints = Math.Round(Bounds.Width, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            return ttMeasurer.WrapMultipleTextFragmentsToTextLines(fragments, fonts, maxWidthPoints);
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
