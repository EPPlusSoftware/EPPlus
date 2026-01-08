using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal class ParagraphContainer : RenderItem
    {
        internal List<FontWrapContainer> Runs = new List<FontWrapContainer>();
        ITextMeasurerWrap measurer;
        TextFragmentCollection _textFragments;
        List<string> _paragraphLines = new List<string>();
        List<string> _textRunContent = new List<string>();

        public override RenderItemType Type => RenderItemType.Text;

        public ParagraphContainer() : base()
        {

        }

        public ParagraphContainer(ExcelDrawingParagraph paragraph, BoundingBox parent) : base()
        {
            measurer = paragraph._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;

            Bounds.Parent = parent;
            Bounds.Width = parent.Width - paragraph.RightMargin;

            InitializeLinesAndRichText(paragraph);
        }

        public ParagraphContainer(BoundingBox parent)
        {
            Bounds.Parent = parent;
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
            var ttMeasurer = (FontMeasurerTrueType)measurer;

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
            return ttMeasurer.WrapMultipleTextFragments(_textFragments, fonts, Bounds.Width);
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
                //Using Regex to avoid empty lines in windows with paragraph.Text.Split(new [] '\r' '\n')
                _paragraphLines = Regex.Split(paragraph.Text, "\r\n|\r|\n").ToList();
                //Gets each actual text run, including linebreak symbols
                foreach (var run in paragraph.TextRuns)
                {
                    _textRunContent.Add(run.Text);
                }
            }
        }

        public override void Render(StringBuilder sb)
        {
            throw new NotImplementedException();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            throw new NotImplementedException();
        }
    }
}
