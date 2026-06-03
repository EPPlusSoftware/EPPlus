using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;


namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    public class SvgParagraphRenderItem : ParagraphRenderItem
    {
        public SvgParagraphRenderItem(RenderTextBody body, BoundingBox parent, string text, bool setDefaultFont = true) : base(parent, body, text, setDefaultFont)
        {
            ImportStyles();
        }
        public SvgParagraphRenderItem(RenderTextBody textBody, BoundingBox parent, IRichTextFormatSimple rtFormat): base(parent, textBody, rtFormat)
        {
            ImportStyles();
        }

        private void ImportStyles()
        {
            //Import RichText data to each run
            foreach (var run in Runs)
            {
                var textRun = (SvgTextRunRenderItem)run;
                var rtOptions = _layoutSystem.InputFragments[run.OriginalRtIdx].RichTextOptions;
                if (_layoutSystem.InputFragments.Count != 0 && run.OriginalRtIdx != -1 && rtOptions is IRichTextFormatSimple)
                {
                    textRun.ImportRichTextData((IRichTextFormatSimple)rtOptions);
                }
                else
                {
                    //If not use the default for the whole paragraph (potentially user specified)
                    run.ImportFontData(DefaultParagraphFont);
                }
            }
        }

        public override RenderItemType Type => RenderItemType.Paragraph;

        protected override TextRunRenderItem CreateTextRun(BoundingBox parent, string displayText, int origRtIdx)
        {
            return new SvgTextRunRenderItem(parent, displayText, origRtIdx);
        }
    }
}
