using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
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
        internal int TotalLength = 0;

        internal List<double> FontSizes = new();

        internal TextFragmentCollection Fragments;

        public TextParagraph(TextFragmentCollection fragments, List<MeasurementFont> fonts)
        {
            List<OpenTypeFont> openTypeFonts = new List<OpenTypeFont>();
            FontSizes = new List<double>();
            Fragments = fragments;

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
