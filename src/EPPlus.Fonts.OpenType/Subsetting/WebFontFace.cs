/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/19/2026         EPPlus Software AB           Web font subsetting for SVG export
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// One subsetted font to declare with @font-face in web output (SVG/HTML).
    /// </summary>
    public sealed class WebFontFace
    {
        internal WebFontFace(string familyName, string sourceFamily, FontSubFamily subFamily, byte[] fontData)
        {
            FamilyName = familyName;
            SourceFamily = sourceFamily;
            SubFamily = subFamily;
            FontData = fontData;
        }

        /// <summary>
        /// The CSS family name of the subset, e.g. "ABCDEF+Carlito". The prefix is derived from the
        /// font and the code points in the subset, so the name is unique per subset: an inline SVG's
        /// @font-face applies to the whole HTML document. Two identical subsets get the same name and
        /// the same data, which is harmless.
        /// </summary>
        public string FamilyName { get; private set; }

        /// <summary>
        /// The family of the font the subset was made from, e.g. "Carlito".
        /// </summary>
        public string SourceFamily { get; private set; }

        /// <summary>
        /// The subfamily of the font the subset was made from.
        /// </summary>
        public FontSubFamily SubFamily { get; private set; }

        /// <summary>
        /// The CSS font-weight descriptor: 700 for bold subfamilies, otherwise 400. Declaring it keeps
        /// the browser from synthesizing bold on top of glyphs that are already bold.
        /// </summary>
        public int FontWeight
        {
            get
            {
                return SubFamily == FontSubFamily.Bold || SubFamily == FontSubFamily.BoldItalic ? 700 : 400;
            }
        }

        /// <summary>
        /// True for italic subfamilies; the CSS font-style descriptor is then "italic", otherwise "normal".
        /// </summary>
        public bool IsItalic
        {
            get { return SubFamily == FontSubFamily.Italic || SubFamily == FontSubFamily.BoldItalic; }
        }

        /// <summary>
        /// The subset as a TrueType font file.
        /// </summary>
        public byte[] FontData { get; private set; }
    }
}