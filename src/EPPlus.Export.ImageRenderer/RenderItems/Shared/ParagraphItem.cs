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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;
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
            
            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();

                runContents.Add(txtRun.Text);
                fontSizes.Add(runFont.Size);
            }

            _textFragments = new TextFragmentCollection(runContents, fontSizes);
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
                lines = measurer.MeasureAndWrapTextLines(textIfEmpty, p.DefaultRunProperties.GetMeasureFont(), maxWidth);
            }
            else
            {
                lines = WrapToSimpleTextLines(p, _textFragments);
            }

            foreach (var line in lines)
            {
                double prevWidth = 0;

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

                foreach (var rtFragment in line.RtFragments)
                {
                    var displayText = line.GetFragmentText(rtFragment);

                    if(p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
                    {
                        AddText(displayText, p.DefaultRunProperties);
                    }
                    else
                    {
                        AddRenderItemTextRun(p.TextRuns[rtFragment.Fragidx], displayText, prevWidth);
                    }

                    TextRunItem runItem = Runs.Last();
                    runItem.YPosition = runLineSpacing;

                    runItem.Bounds.Width = rtFragment.Width;
                    prevWidth += rtFragment.Width;
                }

                lastDescent = line.LargestDescent;
            }
            Bounds.Height = runLineSpacing + lastDescent;
            Bounds.Width = greatestWidth;
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

            var maxWidthPoints = Math.Round(ParentTextBody.MaxWidth, 0, MidpointRounding.AwayFromZero).PixelToPoint();
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
