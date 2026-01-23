using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
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
    /// If wrapping is true, uses parent rects width as maximum
    /// </summary>
    internal class FontWrapContainer : TextContainerBase
    {
        protected FontMeasurerTrueType _measurer;

        /// <summary>
        /// If this is true it is presumed there is also a parent with a maxwidth
        /// If there is no MaxWidth this class will not wrap
        /// </summary>
        public bool WrapText = false;

        public double MaxWidthPixels
        { 
            get 
            { 
                if(Parent != null)
                {
                    return Parent.Width;
                }
                else
                {
                    return double.NaN;
                }
            } 
        }

        public FontWrapContainer(FontMeasurerTrueType txtMeasurer, bool initDefaults = true) : base(initDefaults)
        {
            Initialize(txtMeasurer);
        }

        public FontWrapContainer(FontMeasurerTrueType txtMeasurer, string content, bool initDefaults = false) : base(content, initDefaults)
        {
            Initialize(txtMeasurer);
        }

        private void Initialize(FontMeasurerTrueType txtMeasurer)
        {
            _measurer = txtMeasurer;
            //Transform = parent.Transform;
            //if (parent != null)
            //{
            //    SetParent(parent);
            //}
        }

        //void SetParent(Rect parent)
        //{
        //    TopDrawingHandler = parent;
        //    Transform.TopDrawingHandler = parent.Transform;
        //}

        private void SplitContentToLines()
        {
            if (Content.Length == 1)
            {
                //If width is NaN textWrapper only applies line endings within the text itself
                var inputWidth = WrapText ? MaxWidthPixels : double.NaN;

                Content = TextWrapper.GetLines(Content[0], _measurer, inputWidth).ToArray();
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
            return TextWrapper.GetContentWidths(Content,_measurer, MaxWidthPixels).ToArray();
        }
    }
}
