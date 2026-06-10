using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.Textbox;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Integration.RichText;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    /// <summary>
    /// Text alignment
    /// </summary>
    public enum TextAlignment
    {
        /// <summary>
        /// Left alignment
        /// </summary>
        Left,
        /// <summary>
        /// Center alignment
        /// </summary>
        Center,
        /// <summary>
        /// Right alignment
        /// </summary>
        Right,
        /// <summary>
        /// Distributes the text words across an entire text line
        /// </summary>
        Distributed,
        /// <summary>
        /// Align text so that it is justified across the whole line.
        /// </summary>
        Justified,
        /// <summary>
        /// Aligns the text with an adjusted kashida length for Arabic text
        /// </summary>
        JustifiedLow,
        /// <summary>
        /// Distributes Thai text specially, specially, because each character is treated as a word
        /// </summary>
        ThaiDistributed
    }

    public enum TextLineSpacing
    {
        /// <summary>
        /// Single line spacing
        /// </summary>
        Single,
        /// <summary>
        /// 1.5 lines
        /// </summary>
        OneAndAHalf,
        /// <summary>
        /// Double line spacing
        /// </summary>
        Double,
        /// <summary>
        /// Exact point spacing
        /// </summary>
        Exactly,
        /// <summary>
        /// Multiple line spacing
        /// </summary>
        Multiple
    }

    public abstract class ParagraphRenderItem : RenderItem
    {
        protected double LeftMargin { get; set; }
        protected double RightMargin { get; set; }        
        protected double LineSpacingAscendantOnly { get; set; }
        protected bool IsFirstParagraph { get; set; }

        protected LayoutSystem _layoutSystem;

        protected RichTextCollectionBase _textFragments = new RichTextCollectionBase();

        public double ParagraphLineSpacing { get; protected set; }

        protected TextAlignment _alignment;

        //After setting alignment we must re-calculate the rows
        public TextAlignment HorizontalAlignment { get { return _alignment; } set { _alignment = value; WrapTextFragmentsAndGenerateTextRuns(); } }
        public List<TextRunRenderItem> Runs { get; set; } = new List<TextRunRenderItem>();
        public TextLineCollection Lines { get; protected set; }
        public bool DisplayBounds { get; set; } = false;

        public override RenderItemType Type => RenderItemType.Paragraph;

        public bool AutoSize = false;

        public FontFormatBase DefaultParagraphFont;

        protected double ParentMaxWidth;
        protected double ParentMaxHeight;

        protected RenderTextBody ParentTextBody { get; set; }

        protected double? _lsMultiplier = null;

        protected bool TextIfEmptyIsNull { get; set; }

        protected bool LinespacingIsExact
        {
            get
            {
                return _lsMultiplier.HasValue == false;
            }
        }

        protected TextLineSpacing _lsType;
        protected double? _centerAdjustment;

        protected ParagraphRenderItem(BoundingBox parent, bool setFallbackDefaultFont = true) : base(parent)
        {
            Bounds.Name = "Paragraph";
            if (setFallbackDefaultFont)
            {
                var defaultFont = new MeasurementFont { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };
                DefaultParagraphFont = new FontFormatBase(defaultFont);
                FillColor = "black";
            }
        }

        protected ParagraphRenderItem(BoundingBox parent, RenderTextBody textBody, bool setFallbackDefaultFont = true) : this(parent, setFallbackDefaultFont)
        {
            InitBasedOnParent(textBody);
            Bounds.Name = "Paragraph";
        }

        protected ParagraphRenderItem(BoundingBox parent, RenderTextBody textBody, string text, bool setFallbackDefaultFont = true) : this(parent, textBody, setFallbackDefaultFont)
        {
            _lsMultiplier = 1d;
            ImportLinesAndTextRunsBase(text);
        }

        protected ParagraphRenderItem(BoundingBox parent, RenderTextBody textBody, IRichTextFormatSimple rtFormat) : this(parent, textBody, false)
        {
            _lsMultiplier = 1d;
            DefaultParagraphFont = new FontFormatBase(rtFormat.Family, rtFormat.SubFamily, rtFormat.Size);
            AddRichText(rtFormat);
        }

        protected ParagraphRenderItem(BoundingBox parent, RenderTextBody textBody, IRichTextFormatDrawing rtFormat) : this(parent, textBody, false)
        {
            AddRichText(rtFormat);
        }

        protected double GetAlignmentHorizontal(TextAlignment txAlignment)
        {
            double x = 0;
            switch (txAlignment)
            {
                case TextAlignment.Left:
                default:
                    x = Bounds.Left + LeftMargin;
                    break;
                case TextAlignment.Center:
                    x = (Bounds.Right / 2) + LeftMargin - RightMargin;
                    break;
                case TextAlignment.Right:
                    x = Bounds.Right - RightMargin;
                    break;
            }

            return x;
        }


        void InitBasedOnParent(RenderTextBody textBody)
        {
            ParentTextBody = textBody;
            ParentMaxWidth = textBody.MaxWidth;
            ParentMaxHeight = textBody.MaxHeight;
            AutoSize = textBody.AutoSize;

            if (AutoSize == false)
            {
                Bounds.Width = textBody.Width;
                Bounds.Height = textBody.Height;
            }
            else
            {
                //Set to max until measured
                Bounds.Width = ParentMaxWidth;
                Bounds.Height = ParentMaxHeight;
            }
        }

        TextLineCollection WrapFragmentsToLines(List<ITextFragmentBase>? fragments = null)
        {
            //This is highly innefficent. Really, LayoutSystem should be 
            //Holding the fragments from the start/wrapping should only be done when textFragments are fully complete
            _layoutSystem = new LayoutSystem(_textFragments);

            //if (fragments == null && _layoutSystem == null)
            //{
            //    _layoutSystem = new LayoutSystem(_textFragments);
            //}

            double maxWidthInPoints;
            if(AutoSize)
            {
                maxWidthInPoints = Math.Round(ParentMaxWidth - RightMargin - LeftMargin, 0, MidpointRounding.AwayFromZero);
            }
            else
            {
                maxWidthInPoints = Bounds.Width;
            }
            return _layoutSystem.Wrap(maxWidthInPoints);
        }

        private void AddRichTextBase(IRichTextFormatSimple rt)
        {
            if (_textFragments == null)
            {
                _textFragments = new RichTextCollectionBase();
            }

            if (string.IsNullOrEmpty(rt.Text) == false)
            {
                _textFragments.Add(rt);
            }
        }

        protected void AddDefaultTextFragment(string text)
        {
            var defaults = new RichTextFormatSimple();
            defaults.Text = text;
            defaults.SetFont(DefaultParagraphFont);

            AddRichTextBase(defaults);
        }

        protected void ImportLinesAndTextRunsBase(string textIfEmpty)
        {
            if(string.IsNullOrEmpty(textIfEmpty))
            {
                TextIfEmptyIsNull = true;
            }
            else
            {
                TextIfEmptyIsNull = false;
            }

            AddDefaultTextFragment(textIfEmpty);
            WrapTextFragmentsAndGenerateTextRuns();
        }

        protected void WrapTextFragmentsAndGenerateTextRuns()
        {
            Lines = WrapFragmentsToLines();
            Runs.Clear();

            //In points
            double widthOfLargestLine = 0;
            //Set to 0 then grow to size of content after wrap/measure
            //This as an empty paragraph should have no real size
            double combinedHeight = 0;

            //has value if there is linespacing otherwise isNaN
            //Don't do this on the actual property as a paragraph can have a fallback linespacing without it being applied
            //(e.g. paragraph linespacing is set in the ooxml but the paragraph contains no textruns)
            //We should not change the 'ParagraphLineSpacing' variable directly here
            double lineSpacingResult = LinespacingIsExact ? ParagraphLineSpacing : double.NaN;

            if (Lines != null && Lines.Count != 0)
            {
                widthOfLargestLine = Lines.LargestWidthWithoutSpace;
                combinedHeight = Lines.GetHeightOfCollection(_lsMultiplier, lineSpacingResult);

                if(AutoSize)
                {
                    Bounds.Width = widthOfLargestLine + RightMargin;
                }
                //SetHorizontalAlignment(widthOfLargestLine);

                int lineIdx = 0;
                foreach (var line in Lines)
                {
                    double lineDist = widthOfLargestLine - line.GetWidthWithoutTrailingSpaces();
                    double prevWidth = CalculatePrevWidthBasedOnAlignment(lineDist);

                    foreach (var lineFragment in line.LineFragments)
                    {
                        var displayText = lineFragment.Text;

                        int rtIdx = -1;
                        if (_layoutSystem != null && lineFragment.OriginalTextFragment != null && _layoutSystem.InputFragments.Count > 0)
                        {
                            rtIdx = _layoutSystem.InputFragments.IndexOf(lineFragment.OriginalTextFragment);
                        }

                        var run = CreateTextRun(Bounds, displayText, rtIdx);
                        //Potentially we could import styling here instead but that leads to multiple issues.
                        //We may need to move it back here for auto-size reasons

                        run.YPosition = Lines.GetBaseLinePosition(lineIdx, lineSpacingResult);
                        run.Bounds.Left = prevWidth;
                        run.Bounds.Width = lineFragment.Width;
                        prevWidth += lineFragment.Width;

                        Runs.Add(run);
                    }
                    lineIdx++;
                }
            }
            Bounds.Height = combinedHeight;
        }

        protected double CalculatePrevWidthBasedOnAlignment(double lineDist)
        {
            double prevWidth = 0;
            if (HorizontalAlignment == TextAlignment.Center)
            {
                //Calculate difference in widths and split to get offset between leftmost position and current line
                prevWidth = lineDist / 2;
            }
            else if (HorizontalAlignment == TextAlignment.Right)
            {
                //Note that the actual bounds with the space will be outside max bounds.
                //This appears to be how excel does it
                prevWidth = lineDist;
            }

            return prevWidth;
        }

        protected void SetHorizontalAlignment(double widthOfLargestLine)
        {
            if (HorizontalAlignment == TextAlignment.Center)
            {
                //Bounds of the paragraph should be bounds of the text itself.
                //Therefore we must know the starting point to set accurate left and offset from left.
                Bounds.Left = GetAlignmentHorizontal(HorizontalAlignment) - (widthOfLargestLine / 2);
            }
            else
            {
                //Bounds of the paragraph should be bounds of the text itself.
                //Therefore we must know the starting point to set accurate left and offset from left.
                Bounds.Left = 0;
            }
        }

        void ImportStyles()
        {
            foreach (var run in Runs)
            {
                var rt = _layoutSystem.InputFragments[run.OriginalRtIdx].RichTextOptions;
                if(rt is RichTextFormatDrawing)
                {
                    //Import drawing data
                    run.ImportRichTextData((RichTextFormatDrawing)rt);
                }
                else if(rt is IRichTextFormatSimple)
                {
                    //Import basic/cell data
                    run.ImportRichTextData((IRichTextFormatSimple)rt);
                }
                else
                {
                    //Import only essential font data
                    run.ImportFontData((IFontFormatBase)rt);
                }
            }
        }

        public void AddRichText(IRichTextFormatSimple richText)
        {
            //TODO: Fix superScript/subScript should apply baseLine changes appropriately

            AddRichTextBase(richText);
            WrapTextFragmentsAndGenerateTextRuns();
        }

        public void AddRichText(IRichTextFormatDrawing richText)
        {
            //adjust size in accordance with baseline
            richText.Size = richText.Baseline == 0 ? richText.Size : (float)(richText.Size * (1 - (Math.Abs(richText.Baseline) / 100)));

            AddRichTextBase(richText);
            WrapTextFragmentsAndGenerateTextRuns();
            ImportStyles();
        }

        public void AddText(string text)
        {
            ImportLinesAndTextRunsBase(text);
            ImportStyles();
        }

        protected abstract TextRunRenderItem CreateTextRun(BoundingBox parent, string displayText, int origRtIdx);
    }
}
