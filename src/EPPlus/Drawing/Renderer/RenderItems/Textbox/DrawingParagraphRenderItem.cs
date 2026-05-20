using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    internal class DrawingParagraphRenderItem : ParagraphRenderItem
    {
        eDrawingTextLineSpacing _lsType;
        double? _lsMultiplier = null;
        List<TextFragment> _newTextFragments;
        int _manualFragmentsStartIndex = -1;
        List<TextFragment> _manualFragments;
        internal DrawingTextbody ParentTextBody { get; set; }
        internal bool DisplayBounds { get; set; } = false;
        private List<TextLineSimple> _lines;
        //Start temp workaround vars
        string _textIfEmpty = null;
        ExcelDrawingParagraph Paragraph { get; set; } = null;
        //end temp workaround vars

        private double? _centerAdjustment = null;

        internal List<double> SpaceWidthsPerLine = new List<double>();

        bool LinespacingIsExact 
        { 
            get
            { 
                return _lsMultiplier.HasValue == false; 
            } 
        }

        public override RenderItemType Type => RenderItemType.Paragraph;

        public DrawingParagraphRenderItem(DrawingTextbody textBody, BoundingBox parent) : base(parent)
        {
            ParentTextBody = textBody;
            Bounds.Name = "Paragraph";
            var defaultFont = new MeasurementFont { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };
            ParagraphFont = defaultFont;

            Layout = OpenTypeFonts.GetTextLayoutEngineForFont(defaultFont);
            ParagraphLineSpacing = GetParagraphLineSpacingInPoints(100, (TextShaper)OpenTypeFonts.GetShaperForFont(defaultFont), defaultFont.Size);
        }

        public DrawingParagraphRenderItem(DrawingTextbody textBody, BoundingBox parent, ExcelDrawingParagraph p, string textIfEmpty = null) : base(parent)
        {
            ParentTextBody = textBody;
            IsFirstParagraph = p == p._paragraphs[0];

            if (p.DefaultRunProperties.Fill != null && p.DefaultRunProperties.Fill.IsEmpty == false)
            {
                if (IsFirstParagraph)
                {
                    if (p.DefaultRunProperties.Fill != null)
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, null);
                    }
                }
                else
                {
                    //Drawingproperties has fallback to firstDefault but excel does not display it so we should not either.
                    if (p.DefaultRunProperties != p._paragraphs.FirstDefaultRunProperties)
                    {
                        this.SetDrawingPropertiesFill(textBody.Theme, p.DefaultRunProperties.Fill, null);
                    }
                    else
                    {
                        var fc = ColorConverter.GetThemeColor(textBody.Theme.ColorScheme.Light1);
                        fc = ColorConverter.GetAdjustedColor(PathFillMode.Norm, fc);
                        FillColor = "#" + fc.ToArgb().ToString("x8").Substring(2);
                        //Use shape fill somehow
                        //Maybe use a name property for fallback theme accent1 color?
                    }
                }
            }
            else
            {
                if (p._paragraphs.FirstDefaultRunProperties != null && p._paragraphs.FirstDefaultRunProperties.Fill != null && p._paragraphs.FirstDefaultRunProperties.Fill.IsEmpty == false)
                {
                    var fill = p._paragraphs.FirstDefaultRunProperties.Fill;
                    this.SetDrawingPropertiesFill(textBody.Theme, fill, null);
                }
            }

            //---Initialize Bounds / Margins-- -
            var indent = 48 * p.IndentLevel;
            LeftMargin = p.LeftMargin + p.Indent + indent;
            RightMargin = p.RightMargin;

            LeftMargin = LeftMargin.PixelToPoint();
            RightMargin = RightMargin.PixelToPoint();

            HorizontalAlignment = (TextAlignment)p.HorizontalAlignment;
            LeftMargin = LeftMargin.PixelToPoint();
            RightMargin = RightMargin.PixelToPoint();

            if (ParentTextBody.AutoSize == false)
            {
                Bounds.Left = 0;
                Bounds.Width = ParentTextBody.MaxWidth;

                //Left is equal to left Paragraph margin
                //Textbody or Textbox are assumed to handle shape/chart margins
                //Paragraph handles only indentations/margins that is applied ON TOP of those margins
                //Paragraph left is the exact position where the text itself starts on the left
                Bounds.Left = GetAlignmentHorizontal(TextAlignment.Left);
                if (HorizontalAlignment == TextAlignment.Center)
                {
                    //Center is a bit strange the bounds really are the same as left or right aligned
                    //It doesn't truly matter as only left min and right max play a role
                    _centerAdjustment = GetAlignmentHorizontal(HorizontalAlignment);
                }
                Bounds.Width = parent.Width - RightMargin - LeftMargin;
            }

            //---Initialize / calculate lines and runs---
            //measurer must be set before AddLinesAndRichText
            ParagraphFont = p.DefaultRunProperties.GetMeasureFont();

            //---Get measurer---
            Layout = OpenTypeFonts.GetTextLayoutEngineForFont(ParagraphFont);

            //---Calculate linespacing---
            int numLines = ParagraphLines.Count;
            _lsType = p.LineSpacing.LineSpacingType;
            ParagraphLineSpacing = GetParagraphLineSpacingInPoints(p.LineSpacing.Value, 
                (TextShaper) OpenTypeFonts.GetShaperForFont(ParagraphFont), 
                ParagraphFont.Size);


            ImportLinesAndTextRuns(p, textIfEmpty);
        }

        private double GetParagraphLineSpacingInPoints(double spacingValue, TextShaper fmExact, float fontSize)
        {
            if (_lsType == eDrawingTextLineSpacing.Exactly)
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

        public void AddOwnText(string text)
        {
            var fragment = new TextFragment();
            fragment.Text = text;
            fragment.Font = ParagraphFont;
            _manualFragments.Add(fragment);

            //if(_newTextFragments == null)
            //{
            //    //This should probably never happen
            //    throw new InvalidOperationException("Must GENERATE textfragments first in the constructor");
            //    //GenerateTextFragments(text);
            //}
            //else
            //{
            //    var fragment = new TextFragment();
            //    fragment.Text = text;
            //    fragment.Font = ParagraphFont;
            //    //_newTextFragments.Add(fragment);
            //    _manualFragments.Add(fragment);
            //}

            //Redo whole thing for now.
            //Import and wrapping really should be completely seperated but can't refactor all of it yet
            //AddTextLinesAndSpacing(Paragraph, _textIfEmpty);
        }

        public void AddOwnText(TextFragment fragment)
        {
            //_manualFragments.Add(fragment);

            if (_newTextFragments == null)
            {
                //This should probably never happen
                throw new InvalidOperationException("Must GENERATE textfragments first in the constructor");
                //GenerateTextFragments(text);
            }
            else
            {
                if(_manualFragmentsStartIndex == -1)
                {
                    _manualFragmentsStartIndex = _newTextFragments.Count;
                }
                //_newTextFragments.Add(fragment);
                _newTextFragments.Add(fragment);
            }


            //Redo whole thing for now.
            //Import and wrapping really should be completely seperated but can't refactor all of it yet
            AddTextLinesAndSpacing(Paragraph, _textIfEmpty);
        }


        internal protected TextRunRenderItem AddRenderItemTextRun(ExcelParagraphTextRunBase origTxtRun, string displayText)
        {
            var targetTxtRun = CreateTextRun(origTxtRun, Bounds, displayText);

            Runs.Add(targetTxtRun);
            return targetTxtRun;
        }

        private void AddText(string text, MeasurementFont font)
        {
            var container = CreateTextRun(font, Bounds, text);
            Runs.Add(container);

            container.Bounds.Name = $"Container{Runs.Count}";
        }

        private void AddText(string text, ExcelTextFont font)
        {
            var mf = font.GetMeasureFont();
            var measurer = OpenTypeFonts.GetTextLayoutEngineForFont(mf);

            var container = CreateTextRun(text, font, Bounds, text);
            Runs.Add(container);
            //Bounds.Width = container.Bounds.Width + 0.001; //TODO: fix for equal width issue
            container.Bounds.Name = $"Container{Runs.Count}";
        }

        void GenerateTextFragments(string text)
        {
            _newTextFragments = new List<TextFragment>();

            if (string.IsNullOrEmpty(text) == false)
            {
                var currentFrag = new TextFragment() { Text = text, Font = ParagraphFont};
                _newTextFragments.Add(currentFrag);
            }
        }

        /// <summary>
        /// Log linebreak positions and sizes of the runs
        /// So that we can easily know what textfragment is on what line and what size it has later
        /// </summary>
        /// <param name="runs"></param>
        void GenerateTextFragments(ExcelDrawingTextRunCollection runs/*, string textIfEmpty*/)
        {
            List<string> runContents = new List<string>();
            List<MeasurementFont> fonts = new List<MeasurementFont>();
            
            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();

                fonts.Add(runFont);
                runContents.Add(txtRun.Text);
            }

            _newTextFragments = new List<TextFragment>();

            for (int i = 0; i < runContents.Count(); i++)
            {
                if (string.IsNullOrEmpty(runContents[i]) == false)
                {
                    var currentFrag = new TextFragment() { Text = runContents[i], Font = fonts[i] };
                    _newTextFragments.Add(currentFrag);
                }
            }
        }

        internal void ImportLinesAndTextRunsDefault(string textIfEmpty)
        {
            GenerateTextFragments(textIfEmpty);

            Bounds.Left = GetAlignmentHorizontal(HorizontalAlignment);
            if (HorizontalAlignment == TextAlignment.Center)
            {
                _centerAdjustment = GetAlignmentHorizontal(HorizontalAlignment);
            }

            AddTextLinesAndSpacing(null, textIfEmpty);
        }
        private void ImportLinesAndTextRuns(ExcelDrawingParagraph p, string textIfEmpty)
        {
            if (p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
            {
                ImportLinesAndTextRunsDefault(textIfEmpty);
            }
            else
            {
                //Log line positions and run sizes
                GenerateTextFragments(p.TextRuns);
                AddTextLinesAndSpacing(p, textIfEmpty);
            }
        }

        private void AddTextLinesAndSpacing(ExcelDrawingParagraph p, string textIfEmpty)
        {
            //Temp workaround
            if (Paragraph == null)
            {
                Paragraph = p;
            }
            if (textIfEmpty == null)
            {
                _textIfEmpty = textIfEmpty;
            }

            _lines = WrapFragmentsToLines();

            //In points
            double lastDescent = 0;
            double lineTop = 0;
            double greatestWidth = 0;

            if (_lines != null && _lines.Count != 0)
            {
                //This could be moved into a textLines collection class
                //START
                var idxOfLargestLine = 0;
                double widthOfLargestLine = _lines[0].GetWidthWithoutTrailingSpaces();

                for (int i = 1; i < _lines.Count; i++)
                {
                    if (_lines[i].Width > widthOfLargestLine)
                    {
                        var ctrLineWidth = _lines[i].GetWidthWithoutTrailingSpaces();
                        SpaceWidthsPerLine.Add(_lines[i].lastFontSpaceWidth);
                        widthOfLargestLine = ctrLineWidth;
                        idxOfLargestLine = i;
                    }
                }
                //END


                if (HorizontalAlignment == TextAlignment.Center && ParentTextBody.AutoSize && _centerAdjustment != null && string.IsNullOrEmpty(textIfEmpty))
                {
                    //Bounds of the paragraph should be bounds of the text itself.
                    //Therefore we must know the starting point to set accurate left and offset from left.
                    Bounds.Left = _centerAdjustment.Value - (widthOfLargestLine / 2);
                }
                else
                {
                    Bounds.Left = 0;
                }
                //if (ParentTextBody.AutoSize)
                //{
                //    //Bounds of the paragraph should be bounds of the text itself.
                //    //Therefore we must know the starting point to set accurate left and offset from left.
                //    Bounds.Left = 0;
                //}

                foreach (var line in _lines)
                {
                    double prevWidth = 0;

                    if (HorizontalAlignment == TextAlignment.Center)
                    {
                        var ctrLineWidth = line.GetWidthWithoutTrailingSpaces();
                        //Calculate difference in widths and split to get offset between leftmost position and current line
                        prevWidth = (widthOfLargestLine - ctrLineWidth) / 2;
                    }
                    else if (HorizontalAlignment == TextAlignment.Right)
                    {
                        //Note that the actual bounds with the space will be outside max bounds.
                        //This appears to be how excel does it
                        var ctrLineWidth = line.GetWidthWithoutTrailingSpaces();
                        prevWidth = widthOfLargestLine - ctrLineWidth;
                    }

                    if (LinespacingIsExact == false)
                    {
                        lineTop += line.LargestAscent + lastDescent;
                    }
                    else
                    {
                        lineTop += ParagraphLineSpacing;
                    }
                    if (line.GetWidthWithoutTrailingSpaces() > greatestWidth)
                    {
                        greatestWidth = line.GetWidthWithoutTrailingSpaces();
                    }

                    foreach (var lineFragment in line.LineFragments)
                    {
                        var displayText = lineFragment.Text;


                        if (p != null && p.TextRuns.Count == 0 && string.IsNullOrEmpty(textIfEmpty) == false)
                        {
                            //Import fallback text with paragraph settings
                            AddText(displayText, p.DefaultRunProperties);
                        }
                        else if (p != null && p.TextRuns.Count != 0)
                        {
                            var rtIdx = _newTextFragments.IndexOf(lineFragment.OriginalTextFragment);
                            if (rtIdx > p.TextRuns.Count - 1)
                            {
                                AddText(displayText, _newTextFragments[rtIdx].Font);
                            }
                            else
                            {
                                //Import Paragraph text run
                                var idx = _newTextFragments.IndexOf(lineFragment.OriginalTextFragment);
                                AddRenderItemTextRun(p.TextRuns[idx], displayText);
                            }
                        }
                        else
                        {
                            //Import fallback with default settings from constructor
                            AddText(displayText, ParagraphFont);
                        }
                        DrawingTextRunRenderItem runItem = (DrawingTextRunRenderItem)Runs.Last();
                        runItem.Bounds.Left = prevWidth;
                        runItem.YPosition = lineTop;

                        runItem.Bounds.Width = lineFragment.Width;
                        prevWidth += lineFragment.Width;
                    }
                    lastDescent = line.LargestDescent;
                }
            }
            Bounds.Height = lineTop + lastDescent;
            Bounds.Width = greatestWidth;
        }

        List<TextLineSimple> WrapFragmentsToLines(List<TextFragment> fragments = null)
        {
            if(fragments == null )
            {
                fragments = _newTextFragments;
            }

            if (fragments.Count > 0)
            {
                if(Layout == null)
                {
                    Layout = OpenTypeFonts.GetTextLayoutEngineForFont((fragments[0].Font));
                }

                var maxWidthPoints = Math.Round(ParentTextBody.MaxWidth, 0, MidpointRounding.AwayFromZero);

                _lines = Layout.WrapRichTextLines(fragments, maxWidthPoints);
                return _lines;
            }
            return new List<TextLineSimple>();
        }

        internal double GetAlignmentHorizontal(TextAlignment txAlignment)
       {
            var area = Bounds;
            double x = 0;
            switch (txAlignment)
            {
                case TextAlignment.Left:
                default:
                    x = area.Left + LeftMargin;
                    break;
                case TextAlignment.Center:
                    x = (area.Right / 2) + LeftMargin - RightMargin;
                    break;
                case TextAlignment.Right:
                    x = area.Right - RightMargin;
                    break;
            }

            return x;
        }
        internal TextRunRenderItem CreateTextRun(ExcelParagraphTextRunBase run, BoundingBox parent, string displayText)
        {
            return new DrawingTextRunRenderItem(parent, run, displayText);
        }
        internal TextRunRenderItem CreateTextRun(string text, ExcelTextFont font, BoundingBox parent, string displayText)
        {
            return new DrawingTextRunRenderItem(parent, text, font, displayText);
        }

        internal TextRunRenderItem CreateTextRun(MeasurementFont font, BoundingBox parent, string displayText)
        {
            return new DrawingTextRunRenderItem(parent, font, displayText);
        }

    }
}
