using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer
{
    internal class TextParagraph
    {
        //prevent creating multiple OpenTypeFonts via cache/indexing
        internal Dictionary<double, OpenTypeFont> FontIndexDict = new();
        internal Dictionary<double, GlyphMappings> GlyphMappings = new();
        internal Dictionary<int, CharInfo> CharLookup = new();
        internal int TotalLength = 0;

        internal List<double> FontSizes = new();
        internal List<string> TextFragments = new();

        internal string AllText;
        internal List<int> AllTextNewLineIndicies = new();
        List<TextFragment> fragmentItems = new List<TextFragment>();

        public TextParagraph(List<string> textFragments, List<MeasurementFont> fonts)
        {
            List<OpenTypeFont> openTypeFonts = new List<OpenTypeFont>();
            FontSizes = new List<double>();
            TextFragments = textFragments;

            var distinctFonts = fonts.Distinct().ToArray();

            //Collect fonts that are actually distinct
            foreach (var distinctFont in distinctFonts)
            {
                var subFont = GetFontSubType(distinctFont.Style);
                var font = GetFont(distinctFont.FontFamily, subFont);
                openTypeFonts.Add(font);
            }

            //Setup lookup for different properties
            for (int i = 0; i < fonts.Count; i++)
            {
                for (int j = 0; j < distinctFonts.Count(); j++)
                {
                    if (fonts[i] == distinctFonts[j])
                    {
                        FontIndexDict.Add(i, openTypeFonts[j]);
                        GlyphMappings.Add(i, openTypeFonts[j].CmapTable.GetPreferredSubtable().GetGlyphMappings());
                    }
                }
                FontSizes.Add(fonts[i].Size);
            }

            //Set data for each fragment and for AllText
            var currentTotalLength = 0;
            for(int i = 0; i < textFragments.Count; i++)
            {
                var currString = textFragments[i];
                var fragment = new TextFragment(currString, currentTotalLength);
                currentTotalLength += currString.Length;
                fragmentItems.Add(fragment);
            }

            AllText = string.Join(string.Empty, textFragments.ToArray());
            //Get the indicies where newlines occur in the combined string
            AllTextNewLineIndicies = GetStartIndicies(AllText);


            //Save minor information about each char so each char knows its line/fragment
            int charCount = 0;
            int lineIndex = 0;

            List<int> currFragments = new();
            List<TextLine> lines = new List<TextLine>();
            int lineStartCharIndex = 0;

            //For each fragment
            for (int i = 0; i < TextFragments.Count; i++)
            {
                var textFragment = TextFragments[i];
                
                //For each char in current fragment
                for (int j = 0; j < textFragment.Length; j++)
                {
                    if (charCount >= AllTextNewLineIndicies[lineIndex])
                    {
                        var line = new TextLine()
                        {   richTextIndicies = currFragments,
                            content = AllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex),
                            startIndex = lineStartCharIndex,
                        };

                        lines.Add(line);
                        currFragments.Clear();
                        lineIndex++;
                    }

                    var info = new CharInfo(charCount, i, lineIndex);
                    CharLookup.Add(charCount, info);
                    charCount++;
                }
                currFragments.Add(i);
            }
        }

        private List<int> GetStartIndicies(string stringsCombined)
        {
            List<int> combinedStartIndicies = new List<int>();

            var strings = stringsCombined.Split([Environment.NewLine], StringSplitOptions.None);
            TotalLength = 0;

            for (int i = 0; i < strings.Count(); i++)
            {
                combinedStartIndicies.Add(strings[i].Length + TotalLength);
                TotalLength += strings[i].Length;
            }

            return combinedStartIndicies;
        }

        public TextParagraph(List<string> textFragment, List<double> fontSizes, Dictionary<double, OpenTypeFont> fontIndexDict) 
        { 

        }

        OpenTypeFont GetFont(string fontName, FontSubFamily subFamily)
        {
            return TextData.GetFontData(fontName, subFamily);
        }

        private FontSubFamily GetFontSubType(MeasurementFontStyles Style)
        {
            if ((Style & (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic)) == (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if ((Style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold)
            {
                return FontSubFamily.Bold;
            }
            else if ((Style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic)
            {
                return FontSubFamily.Italic;
            }

            return FontSubFamily.Regular;
        }



        internal void GetNextFragmentInfo()
        {

        }
    }
}
