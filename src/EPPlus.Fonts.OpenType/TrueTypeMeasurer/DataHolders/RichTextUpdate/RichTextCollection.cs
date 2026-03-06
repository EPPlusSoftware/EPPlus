using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders.RichTextUpdate
{
    internal class RichTextCollection : IEnumerable<RichText>
    {
        internal List<RichText> RichTexts { get; private set; }
        List<int> _fragmentStartIndicies = new List<int>();
        internal MeasurementFont[] DistinctFonts { get; private set; }
        internal Dictionary<int, int> IdxToDistinctFontIndex = new Dictionary<int, int>();

        public RichTextCollection(List<string> textFragments, List<MeasurementFont> fonts)
        {

            if (textFragments.Count() != fonts.Count)
            {
                throw new InvalidOperationException($"RichTextCollection Constructor: richTexts list and fonts list must be equal." +
                    $"Counts:" +
                    $"textFragment: {textFragments.Count()}" +
                    $"fonts: {fonts.Count()}");
            }

            for (int i = 0; i < textFragments.Count; i++)
            {
                var rt = new RichText(textFragments[i], fonts[i]);
                RichTexts.Add(rt);
            }

            //If a lot of fonts are exactly the same its good to store the unique ones in terms of fontFamily and Style
            DistinctFonts = fonts.GroupBy(x=> new { x.FontFamily, x.Style}).Select(g => g.First()).ToArray();

            for (int i = 0; i < RichTexts.Count; i++)
            {
                for (int j = 0; j < DistinctFonts.Count(); j++)
                {
                    if (RichTexts[i].Font == DistinctFonts[j])
                    {
                        IdxToDistinctFontIndex.Add(i, j);
                    }
                }
            }

            InitializeCollection();
        }

        void InitializeCollection()
        {
            _fragmentStartIndicies.Add(0);
            string combinedText = "";
            int currentIndex = 0;

            for (int i = 0; i < RichTexts.Count; i++)
            {
                combinedText += RichTexts[i].Text;
                currentIndex += RichTexts[i].Text.Length;
                _fragmentStartIndicies.Add(currentIndex);
            }
        }

        public IEnumerator<RichText> GetEnumerator()
        {
            return RichTexts.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return RichTexts.GetEnumerator();
        }
    }
}
