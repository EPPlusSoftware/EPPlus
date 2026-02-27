using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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

        eDrawingTextLineSpacing _lsType;
        double _lineSpacingAscendantOnly;
        double? _lsMultiplier = null;
        internal bool IsFirstParagraph { get; private set; }
        List<string> _paragraphLines = new List<string>();
        protected List<string> _textRunDisplayText = new List<string>();

        TextFragmentCollection _textFragments;
        List<EPPlus.Fonts.OpenType.Integration.TextFragment> _newTextFragments;
        internal protected MeasurementFont _paragraphFont;
        internal TextBodyItem ParentTextBody { get; set; }
        internal double ParagraphLineSpacing { get; private set; }
        internal eTextAlignment HorizontalAlignment { get; private set; } = eTextAlignment.Left;
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

            _leftMargin = _leftMargin.PixelToPoint();
            _rightMargin = _rightMargin.PixelToPoint();

            HorizontalAlignment = p.HorizontalAlignment;

            if (ParentTextBody.AutoSize == false)
            {
                Bounds.Left = 0;
                Bounds.Width = ParentTextBody.MaxWidth;

                //Left is equal to left Paragraph margin
                //Textbody or Textbox are assumed to handle shape/chart margins
                //Paragraph handles only indentations/margins that is applied ON TOP of those margins
                //Paragraph left is the exact position where the text itself starts on the left
                Bounds.Left = GetAlignmentHorizontal(HorizontalAlignment);
                Bounds.Width = parent.Width - _rightMargin - _leftMargin;

                //Bounds.Width = ParentTextBody.Width;
                //var globBounds = Bounds.GetGlobalBoundingbox();
                //Bounds.Width = parent.Width;
                //Bounds.Left = GetAlignmentHorizontal(HorizontalAlignment);
                //Bounds.Width = parent.Width - _rightMargin - _leftMargin;
                //Bounds.Width = parent.Width;
                //Bounds.Left = _leftMargin;
                ////if (HorizontalAlignment != eTextAlignment.Center)
                ////{
                ////    Bounds.Left = GetAlignmentHorizontal(HorizontalAlignment);
                ////}
                ////else
                ////{
                ////    //Center is a bit strange the bounds really are the same as left or right aligned
                ////    //It doesn't truly matter as only left min and right max play a role
                ////    Bounds.Left = GetAlignmentHorizontal(eTextAlignment.Left);
                ////}
                //Bounds.Width = parent.Width - _rightMargin - _leftMargin;
            }

            //---Get measurer---
            _measurer = p._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;

            //---Calculate linespacing---
            int numLines = _paragraphLines.Count;
            _lsType = p.LineSpacing.LineSpacingType;
            ParagraphLineSpacing = GetParagraphLineSpacingInPoints(p.LineSpacing.Value, _measurer);

            //---Initialize / calculate lines and runs---
            //measurer must be set before AddLinesAndRichText
            _paragraphFont = p.DefaultRunProperties.GetMeasureFont();
            _measurer.SetFont(_paragraphFont);

            AddLinesAndTextRuns(p, textIfEmpty);
        }

        private double GetParagraphLineSpacingInPoints(double spacingValue, ITextMeasurerWrap fmExact)
        {
            if (_lsType == eDrawingTextLineSpacing.Exactly)
            {
                if (IsFirstParagraph)
                {
                    _lineSpacingAscendantOnly = spacingValue;
                }
                return spacingValue;
            }
            else
            {
                var multiplier = (spacingValue / 100);
                _lsMultiplier = multiplier;
                if (IsFirstParagraph)
                {
                    _lineSpacingAscendantOnly = multiplier * fmExact.GetBaseLine();
                }
                return multiplier * fmExact.GetSingleLineSpacing();
            }
        }

        internal protected TextRunItem AddRenderItemTextRun(ExcelParagraphTextRunBase origTxtRun, string displayText, double startingX)
        {
            var targetTxtRun = CreateTextRun(origTxtRun, Bounds, displayText);
            targetTxtRun.Bounds.Left = startingX;

            Runs.Add(targetTxtRun);
            return targetTxtRun;
        }

        public void AddText(string text, ExcelTextFont font, bool isOld)
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


        public void AddText(string text, ExcelTextFont font)
        {
            var measurer = new FontMeasurerTrueType();
            //var displayText = measurer.MeasureAndWrapTextLines(text, font.GetMeasureFont(), ParentTextBody.MaxWidth);

            var container = CreateTextRun(text, font, Bounds, text);
            //container.BaseLineSpacing = _lineSpacingAscendantOnly;
            //container.LineSpacingPerNewLine = _lsMultiplier.Value * _measurer.GetSingleLineSpacing().PointToPixel(true);
            Runs.Add(container);
            //Bounds.Width = container.Bounds.Width + 0.001; //TODO: fix for equal width issue
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
            List<MeasurementFont> fonts = new List<MeasurementFont>();
            
            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();

                fonts.Add(runFont);
                runContents.Add(txtRun.Text);
                fontSizes.Add(runFont.Size);
            }

            _textFragments = new TextFragmentCollection(runContents, fontSizes);

            _newTextFragments = new List<EPPlus.Fonts.OpenType.Integration.TextFragment>();

            for (int i = 0; i < runContents.Count(); i++)
            {
                if (string.IsNullOrEmpty(runContents[i]) == false)
                {
                    var currentFrag = new TextFragment() { Text = runContents[i], Font = fonts[i] };
                    _newTextFragments.Add(currentFrag);
                }
            }
        }

        private void AddLinesAndTextRuns(ExcelDrawingParagraph p, string textIfEmpty)
        {
            //Log line positions and run sizes
            GenerateTextFragments(p.TextRuns);

            var lines = new List<TextLineSimple>();
            //In points
            double lastDescent = 0;
            bool lineSpacingIsExact = _lsMultiplier.HasValue == false;
            double runLineSpacing = 0;
            double greatestWidth = 0;

            if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
            {
                var measurer = new FontMeasurerTrueType();
                var maxWidth = ParentTextBody.MaxWidth + 0.001; //TODO: fix for equal width issue;
                lines = measurer.MeasureAndWrapTextLines_New(textIfEmpty, p.DefaultRunProperties.GetMeasureFont(), maxWidth);
            }
            else
            {
                lines = WrapToSimpleTextLines(p);
            }

            if (lines != null)
            {
                foreach (var line in lines)
                {
                    double prevWidth = 0;

                    if (HorizontalAlignment == eTextAlignment.Center)
                    {
                        var ctrLineWidth = line.GetWidthWithoutTrailingSpaces();
                        prevWidth = -ctrLineWidth / 2;
                    }

                    if (lineSpacingIsExact == false)
                    {
                        runLineSpacing += line.LargestAscent + lastDescent;
                    }
                    else
                    {
                        runLineSpacing += ParagraphLineSpacing;
                    }
                    if (line.Width > greatestWidth)
                    {
                        greatestWidth = line.Width;
                    }

                    foreach (var lineFragment in line.LineFragments)
                    {
                        var displayText = line.GetLineFragmentText(lineFragment);

                        if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
                        {
                            AddText(displayText, p.DefaultRunProperties);
                        }
                        else
                        {
                            AddRenderItemTextRun(p.TextRuns[lineFragment.RtFragIdx], displayText, prevWidth);
                        }

                        TextRunItem runItem = Runs.Last();
                        runItem.YPosition = runLineSpacing;

                        runItem.Bounds.Width = lineFragment.Width;
                        prevWidth += lineFragment.Width;
                    }

                    lastDescent = line.LargestDescent;
                }
            }
            Bounds.Height = runLineSpacing + lastDescent;
            Bounds.Width = greatestWidth;
            //if (HorizontalAlignment == eTextAlignment.Center)
            //{
            //    Bounds.Left = Bounds.Left - greatestWidth;
            //}
        }

        List<TextLineSimple> WrapToSimpleTextLines(ExcelDrawingParagraph p)
        {
            var ttMeasurer = (FontMeasurerTrueType)_measurer;

            if (_newTextFragments.Count > 0)
            {
                ttMeasurer.SetFont(_newTextFragments[0].Font);
                var maxWidthPoints = Math.Round(ParentTextBody.MaxWidth, 0, MidpointRounding.AwayFromZero);
                return ttMeasurer.WrapMultipleTextFragmentsToTextLines_New(_newTextFragments, maxWidthPoints);
            }
            return new List<TextLineSimple>();
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

            return x;
        }

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
