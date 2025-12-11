using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    /// <summary>
    /// Only measures widths and handles the actual strings
    /// This class only changes its bounding box if initDefaults = true
    /// Any other changes to its size must be determined by the caller
    /// </summary>
    internal class FontTextContainerBase : TextContainerBase
    {
        FontMeasurerTrueType measurer;

        /// <summary>
        /// If this is true it is presumed there is also a maxWidth
        /// If there is no MaxWidth this class will not wrap
        /// </summary>
        public bool WrapText = false;

        public double MaxWidthPixels = double.NaN;

        public FontTextContainerBase(FontMeasurerTrueType txtMeasurer, bool initDefaults = true) : base(initDefaults)
        {
            measurer = txtMeasurer;
        }

        public FontTextContainerBase(FontMeasurerTrueType txtMeasurer, string content, bool initDefaults = false) : base(content, initDefaults)
        {
            measurer = txtMeasurer;
        }

        private void SplitContentToLines()
        {
            if (Content.Length == 1)
            {
                //If width is NaN textWrapper handles it correctly anyhow
                var inputWidth = WrapText ? MaxWidthPixels : double.NaN;

                Content = TextWrapper.GetLines(Content[0], measurer, inputWidth).ToArray();
            }
        }

        internal int GetNumberOfLines()
        {
            SplitContentToLines();
            return Content.Length;
        }

        /// <summary>
        /// Returns lines of content in pixel width
        /// </summary>
        /// <returns></returns>
        public double[] GetContentWidths()
        {
            SplitContentToLines();
            return TextWrapper.GetContentWidths(Content,measurer, MaxWidthPixels).ToArray();
        }
    }
}
