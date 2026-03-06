using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders.RichTextUpdate
{
    internal class TextFragmentPosition
    {
        int start;
        int end;
    }

    internal class TextWrapper
    {
        public TextWrapper(List<string> textFragments) 
        {
            List<int> fragmentStartIndicies = new List<int>(0);
            string combinedText = "";
            int currentIndex = 0;

            for (int i = 0; i < textFragments.Count; i++)
            {
                combinedText += textFragments[i];
                currentIndex += textFragments[i].Length;
                fragmentStartIndicies.Add(currentIndex);
            }
        }
    }
}
