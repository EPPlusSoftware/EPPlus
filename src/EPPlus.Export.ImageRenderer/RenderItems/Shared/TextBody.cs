using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using System.Collections.Generic;


namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal class TextBody : FontWrapContainer
    {
        internal List<FontWrapContainer> Runs = new List<FontWrapContainer>();

        public bool AllowOverflow;

        public TextBody(FontMeasurerTrueType txtMeasurer, BoundingBox parent, bool initDefaults = true) : base(txtMeasurer, initDefaults)
        {
            transform.Parent = parent.transform;
        }

        public void AddText(string text)
        {
            var container = new FontWrapContainer(measurer, true);
            container.transform.Parent = transform;

            Runs.Add(container);

            container.transform.Name = $"Container{Runs.Count}";

            container.SetContent(text);
        }
    }
}
