using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    public interface ITextFragmentBase
    {
        public string Text { get; set; }
        /// <summary>
        /// Store rich-text info.
        /// We must extract font info from this but nothing else is supposed to be done with this within opentype
        /// but we hold the data so users may more easily recognize which rich text this is in the output.
        /// </summary>
        public IRichTextFormatBase RichTextOptions { get; set; }
        public ShapingOptions Options { get; set; }
        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }

        public abstract float Size { get;}
    }
}
