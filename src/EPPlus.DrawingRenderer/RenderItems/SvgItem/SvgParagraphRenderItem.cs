using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    public class SvgParagraphRenderItem : ParagraphRenderItem
    {
        public SvgParagraphRenderItem(RenderTextBody body, BoundingBox parent, string text, bool setDefaultFont = true) : base(parent, body, text, setDefaultFont)
        {
            ////In svg text is rendered upwards from base-line position
            //if(string.IsNullOrEmpty(text) == false)
            //{
            //    //That means it will not be visible/within bounds unless adjusted
            //    Bounds.Top += Bounds.Height;
            //}
        }

        public override RenderItemType Type => RenderItemType.Paragraph;

        protected override TextRunRenderItem CreateTextRun(BoundingBox parent, string displayText, int origRtIdx)
        {
            return new SvgTextRunRenderItem(parent, displayText, origRtIdx);
        }
    }
}
