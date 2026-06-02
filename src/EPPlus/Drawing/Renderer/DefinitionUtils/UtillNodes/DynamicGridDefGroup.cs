using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils.UtillNodes
{
    internal class DynamicGridDefGroup : RenderItem
    {
        protected const string _urlRef = "\"url(#{0})\"";

        internal string Id { get; private set; }

        internal List<RenderItem> Items = new List<RenderItem>();

        internal LinePattern LnPatternHorizontal { get; private set; }
        internal LinePattern LnPatternVertical { get; private set; }

        FadeOutMask maskItem;

        LinearGradient fadeOutGradient;

        DynamicGridItem gridItem;

        internal string TopId { get; private set; }

        public DynamicGridDefGroup(DrawingBase renderer, string id, int linesX, int linesY): base(renderer) 
        {
            Id = id;

            string lnGradId = id + "_lnGradFade";
            string maskId = id + "_maskFade";
            string horzId = id + "_lnHorizontal";
            string verId = id + "_lnVertical";

            fadeOutGradient = CreateFadeOutGradient(lnGradId);
            Items.Add(fadeOutGradient);

            maskItem = new FadeOutMask(renderer, maskId, string.Format("url(#{0})", lnGradId));

            Items.Add(maskItem);

            LnPatternHorizontal = new LinePattern(renderer, horzId, LinePatternType.Horizontal);
            LnPatternVertical = new LinePattern(renderer, verId, LinePatternType.Vertical);

            SetNumLines(linesX, linesY);

            Items.Add(LnPatternHorizontal);
            Items.Add(LnPatternVertical);

            gridItem = new DynamicGridItem(renderer, id, maskId, horzId, verId);
            Items.Add(gridItem);
        }

        internal void SetNumLines(int numLinesHorizontal, int numLinesVertical)
        {
            LnPatternHorizontal.SetNumberOfLines(numLinesHorizontal);
            LnPatternVertical.SetNumberOfLines(numLinesVertical);
        }

        LinearGradient CreateFadeOutGradient(string id)
        {
            var colors = new List<Color>();
            colors.Add(Color.Black);
            colors.Add(Color.White);

            var stops = new List<double>();
            stops.Add(0);
            stops.Add(100);

            fadeOutGradient = new LinearGradient(DrawingRenderer, id);
            fadeOutGradient.GradientFillExtra = new DrawGradientFill(colors, stops);

            return fadeOutGradient;
        }

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<g>");

            foreach (var item in Items)
            {
                item.Render(sb);
            }

            sb.Append("</g>");
        }
    }
}
