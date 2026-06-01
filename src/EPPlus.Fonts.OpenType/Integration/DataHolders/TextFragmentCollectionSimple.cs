using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Fonts.OpenType.Integration
{
    public class TextFragmentCollectionSimple : List<ITextFragmentBase>, IEnumerable<ITextFragmentBase>
    {
        public TextFragmentCollectionSimple(List<MeasurementFont> fonts, List<string> texts) : base()
        {
            for (int i = 0; i< fonts.Count; i++)
            {
                var tf = new TextFragment() { Text = texts[i] };
                tf.RichTextOptions.SetFont(fonts[i]);
                Add(tf);
            }
        }
    }
}
