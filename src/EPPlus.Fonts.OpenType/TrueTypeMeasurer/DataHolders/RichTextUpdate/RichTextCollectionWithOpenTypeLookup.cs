using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders.RichTextUpdate
{
    internal class RichTextCollectionWithOpenTypeLookup
    {
        internal RichTextCollection Collection;

        internal Dictionary<double, OpenTypeFont> FontIndexDict = new();
        internal Dictionary<double, GlyphMappings> GlyphMappings = new();
        public RichTextCollectionWithOpenTypeLookup(RichTextCollection collection)
        {
            Collection = collection;
            List<OpenTypeFont> openTypeFonts = new List<OpenTypeFont>();

            //Only Get OpenType fonts that are actually distinct
            foreach (var distinctFont in collection.DistinctFonts)
            {
                var subFont = GetFontSubType(distinctFont.Style);
                var font = GetFont(distinctFont.FontFamily, subFont);
                openTypeFonts.Add(font);
            }

            for (int i = 0; i < collection.Count(); i++)
            {
                var distinctIndex = collection.IdxToDistinctFontIndex[i];
                FontIndexDict.Add(i, openTypeFonts[distinctIndex]);
                GlyphMappings.Add(i, openTypeFonts[distinctIndex].CmapTable.GetPreferredSubtable().GetGlyphMappings());
            }
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
    }
}
