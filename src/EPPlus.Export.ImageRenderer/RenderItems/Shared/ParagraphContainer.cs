using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal class ParagraphContainer : RenderItem
    {
        internal List<FontWrapContainer> Runs = new List<FontWrapContainer>();

        public override RenderItemType Type => RenderItemType.Text;

        public ParagraphContainer() : base()
        {

        }

        public ParagraphContainer(ExcelDrawingParagraph paragraph, BoundingBox parent) : base()
        {
            var measurer = paragraph._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
            var ttMeasurer = (FontMeasurerTrueType)measurer;

            for (int i = 0; i < paragraph.TextRuns.Count(); i++)
            {
                var txtRun = paragraph.TextRuns[i];
                var runFont = txtRun.GetMeasureFont();

                ttMeasurer.SetFont(runFont);

                AddText(paragraph.TextRuns[i].Text, ttMeasurer);
            }
            //paragraph.DefaultRunProperties
            //SetDrawingPropertiesFill(paragraph.DefaultRunProperties.Fill, paragraph._prd.s)
            //fill
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
