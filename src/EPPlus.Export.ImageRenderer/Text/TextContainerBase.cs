using System;
using EPPlus.Graphics;

namespace EPPlus.Export.ImageRenderer.Text
{
    /// <summary>
    /// Simple base-class. 
    /// A Rect that also holds an arr of strings
    /// No assumptions are made about where or if the text is placed inside the rect.
    /// </summary>
    internal class TextContainerBase : BoundingBox
    {
        protected string[] Content = null;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="initDefaults">If true initializes the container to 64 width and 20 height </param>
        public TextContainerBase(bool initDefaults = true)
        {
            if (initDefaults)
            {
                //Right and Bottom Pixel defaults for a Cell in excel at 96 PPI
                //(15pts height, 8.43pts width)
                Left = 0; Top = 0; Width = 64; Height = 20d;
                SetContent("Some Text");
            }
        }

        public TextContainerBase(string content, bool initDefaults = true)
        {
            if (initDefaults)
            {
                //Right and Bottom Pixel defaults for a Cell in excel at 96 PPI
                //(15pts height, 8.43pts width)
                Left = 0; Top = 0; Width = 64; Height = 20d;
            }

            SetContent(content);
        }

        public void SetContent(string content)
        {
            Content = new string[] { content };
        }

        public string GetContent()
        {
            var combinedString = "";
            combinedString = string.Join(Environment.NewLine, Content);
            return combinedString;
        }
    }
}
