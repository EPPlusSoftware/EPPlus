using EPPlus.DrawingRenderer;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Integration.RichText;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Drawing;

namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    internal class DrawingParagraphRenderItem : ParagraphRenderItem
    {
        /// <summary>
        /// Create basic empty paragraph
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="parent"></param>
        public DrawingParagraphRenderItem(RenderContext renderContext, DrawingTextBody textBody, BoundingBox parent)
      : base(renderContext, parent, textBody)
        {
            ParagraphLineSpacing = GetParagraphLineSpacingInPoints(100, (TextShaper)RenderContext.FontEngine.GetShaperForFont(DefaultParagraphFont), DefaultParagraphFont.Size);
        }

        /// <summary>
        /// Create paragraph and import a singular text/richText
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="parent"></param>
        /// <param name="text"></param>
        public DrawingParagraphRenderItem(RenderContext renderContext, DrawingTextBody textBody, BoundingBox parent, string text)
           : this(renderContext, textBody, parent)
        {
            ImportLinesAndTextRunsDefault(text);
        }

        /// <summary>
        /// Create paragraph and import all textruns from ExcelDrawingParagraph
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="parent"></param>
        /// <param name="p"></param>
        /// <param name="textIfEmpty"></param>
        public DrawingParagraphRenderItem(RenderContext renderContext, DrawingTextBody textBody, BoundingBox parent, ExcelDrawingParagraph p, string textIfEmpty = null)
            : base(renderContext, parent, textBody, false)
        {
            IsFirstParagraph = p == p._paragraphs[0];
            ImportStyleInfo(textBody, p);
            HorizontalAlignment = (TextAlignment)(int)p.HorizontalAlignment;
            ImportMarginAndIndent(p);

            DefaultParagraphFont = new FontFormatBase(p.DefaultRunProperties.GetMeasureFont());

            ImportLineSpacing(p.LineSpacing.LineSpacingType, p.LineSpacing.Value);

            ImportLinesAndTextRuns(p, textIfEmpty);
        }

        private double GetParagraphLineSpacingInPoints(double spacingValue, TextShaper fmExact, float fontSize)
        {
            if (_lsType == TextLineSpacing.Exactly)
            {
                if (IsFirstParagraph)
                {
                    LineSpacingAscendantOnly = spacingValue;
                }
                return spacingValue;
            }
            else
            {
                var multiplier = (spacingValue / 100);
                _lsMultiplier = multiplier;
                if (IsFirstParagraph)
                {
                    LineSpacingAscendantOnly = multiplier * fmExact.GetAscentInPoints(fontSize);
                }
                return multiplier * fmExact.GetLineHeightInPoints(fontSize);
            }
        }

        private void ImportLinesAndTextRuns(ExcelDrawingParagraph p, string textIfEmpty)
        {
            if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
            {
                ImportLinesAndTextRunsDefault(textIfEmpty, p.DefaultRunProperties);
            }
            else
            {
                //Log line positions and run sizes
                GenerateRichText(p.TextRuns);
                TextIfEmptyIsNull = string.IsNullOrEmpty(textIfEmpty);
                //Initalize and wrap textruns
                WrapTextFragmentsAndGenerateTextRuns();
                ImportStyles(p.TextRuns, p.DefaultRunProperties);
            }
        }

        internal void ImportLinesAndTextRunsDefault(string textIfEmpty, ExcelTextFont font = null)
        {
            ImportLinesAndTextRunsBase(textIfEmpty);
            //Import RichText data to each run
            foreach (var run in Runs)
            {
                var textRun = (DrawingTextRunRenderItem)run;
                ImportStyleFallback(font, textRun);
            }
        }

        /// <summary>
        /// Log linebreak positions and sizes of the runs
        /// So that we can easily know what textfragment is on what line and what size it has later
        /// </summary>
        /// <param name="runs"></param>
        void GenerateRichText(ExcelDrawingTextRunCollection runs/*, List<ShapingOptions>? optionLst = null*/)
        {
            //var lstOfRichText = runs.ExportToOpenTypeFormat();
            var lstOfRichText = runs.ExportToImageRendererFormat();
            foreach (var rt in lstOfRichText)
            {
                _textFragments.Add(rt);
            }
        }

        private void ImportStyleInfo(DrawingTextBody textBody, ExcelDrawingParagraph p)
        {
            ExcelChartStyleColorManager colorManager = null;

            if (textBody._drawing != null)
            {
                if(textBody._drawing.DrawingType is eDrawingType.Shape)
                {
                    colorManager = textBody._drawing.As.Shape.ThemeStyles.FontReference.Color;
                }
                else if(textBody._drawing.DrawingType is eDrawingType.Chart)
                {
                    //Do nothing. Indivdual parts of the chart have different defaults. Let the chart part that created this textbox set its fill instead.
                }
            }

            Color defaultTextColor = textBody.Theme.ColorScheme.GetColorByEnum(eSchemeColor.Text1).GetColor();

            //If this paragraph has defaults of its own enter here
            if (p.DefaultRunProperties.Fill != null && p.DefaultRunProperties.Fill.IsEmpty == false)
            {
                if (IsFirstParagraph)
                {
                    if (p.DefaultRunProperties.Fill != null)
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                }
                else
                {
                    //Drawingproperties has fallback to firstDefault but excel does not display it so we should not either.
                    if (p.DefaultRunProperties != p._paragraphs.FirstDefaultRunProperties)
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                    else
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p._paragraphs.FirstDefaultRunProperties.Fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                }
            }
            else
            {
                if(p._paragraphs.Count != 0)
                {
                    //Fallback to the defaults of the first paragraph
                    if (p._paragraphs[0].DefaultRunProperties != null && p._paragraphs[0].DefaultRunProperties.Fill != null && p._paragraphs[0].DefaultRunProperties.Fill.IsEmpty == false)
                    {
                        var fill = p._paragraphs[0].DefaultRunProperties.Fill;
                        this.SetDrawingPropertiesFill(textBody.Theme, fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                    else if(p.DefaultRunProperties.Fill != null)
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                    else
                    {
                        //Should never happen.
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, colorManager, UserSpaceSettings.ObjectBoundingBox, defaultTextColor);
                    }
                }
            }
        }

        private void ImportMarginAndIndent(ExcelDrawingParagraph p)
        {
            var indent = 48 * p.IndentLevel;
            LeftMargin = p.LeftMargin + p.Indent + indent;
            RightMargin = p.RightMargin;

            LeftMargin = LeftMargin.PixelToPoint();
            RightMargin = RightMargin.PixelToPoint();

            _alignment = (TextAlignment)p.HorizontalAlignment;
            LeftMargin = LeftMargin.PixelToPoint();
            RightMargin = RightMargin.PixelToPoint();
        }

        private void ImportAlignment(bool isAutoSize, double maxWidth, double parentWidth)
        {
            //if (isAutoSize == false)
            //{
                //Bounds.Left = 0;
                //Bounds.Width = ParentMaxWidth;

                ////Left is equal to left Paragraph margin
                ////Textbody or Textbox are assumed to handle shape/chart margins
                ////Paragraph handles only indentations/margins that is applied ON TOP of those margins
                ////Paragraph left is the exact position where the text itself starts on the left
                //Bounds.Left = GetAlignmentHorizontal(TextAlignment.Left);
                //if (HorizontalAlignment == TextAlignment.Center)
                //{
                //    //Center is a bit strange the bounds really are the same as left or right aligned
                //    //It doesn't truly matter as only left min and right max play a role
                //    _centerAdjustment = GetAlignmentHorizontal(HorizontalAlignment);
                //}
                //Bounds.Width = parentWidth - RightMargin - LeftMargin;
            //}
        }

        private void ImportLineSpacing(eDrawingTextLineSpacing lsType, double lineSpacingValue)
        {
            _lsType = (TextLineSpacing)lsType;
            var shaper = (TextShaper)RenderContext.FontEngine.GetShaperForFont(DefaultParagraphFont);

            ParagraphLineSpacing = GetParagraphLineSpacingInPoints(
                lineSpacingValue,
                shaper,
                DefaultParagraphFont.Size);
        }

        void ImportStyleFallback(ExcelTextFont font, DrawingTextRunRenderItem run)
        {
            if (font != null)
            {
                //IF there is a default ExcelTextFont use it
                run.ImportExcelTextFont(font, DefaultParagraphFont);
            }
            else
            {
                //If not use the default for the whole paragraph (potentially user specified)
                run.ImportFontData(DefaultParagraphFont);
            }
        }

        void ImportStyles(ExcelDrawingTextRunCollection textRuns, ExcelTextFont font)
        {
            //Import RichText data to each run
            foreach (var run in Runs)
            {
                var textRun = (DrawingTextRunRenderItem)run;

                if (textRuns.Count != 0 && run.OriginalRtIdx != -1)
                {
                    //Import existing textrun
                    textRun.ImportTextRunBase(textRuns[run.OriginalRtIdx], _layoutSystem.InputFragments[run.OriginalRtIdx].RichTextOptions);
                }
                else
                {
                    //Import default properties or fallback font
                    ImportStyleFallback(font, textRun);
                }
            }
        }
        protected override TextRunRenderItem CreateTextRun(BoundingBox parent, string displayText, int origRtIdx)
        {
            return new DrawingTextRunRenderItem(Bounds, displayText, origRtIdx);
        }
    }
}
