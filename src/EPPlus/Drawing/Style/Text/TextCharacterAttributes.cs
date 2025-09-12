using OfficeOpenXml.Style;
using System;

namespace OfficeOpenXml.Drawing.Style.Text
{
    internal class TextCharacterAttributes
    {
        internal TextCharacterAttributes()
        { 

        }

        /// <summary>
        /// Specifies alt language for UI
        /// </summary>
        internal string AltLang;

        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold;

        /// <summary>
        /// Baseline for superscript and subscript
        /// Based on percentage of the fontsize
        /// </summary>
        public double Baseline;

        /// <summary>
        /// References link targetname for CustomXML link properties
        /// </summary>
        internal string BMK;

        public eTextCapsType Capitalization;

        /// <summary>
        /// Flag for checking spelling grammar etc
        /// Note: Should probably be set whenever user changes the node text in epplus.
        /// </summary>
        internal bool Dirty;

        /// <summary>
        /// Performance improvement. No need to spell-check if already known to be misspelled.
        /// </summary>
        internal bool SpellingError;

        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic;

        double _kerning;

        public double Kerning
        {
            get 
            {
                return _kerning / 100d;
            }
            set 
            {
               if (value < 0 || value > 4000) throw new ArgumentOutOfRangeException("kerning", "Fontsize must be between 0 and 4000");
               _kerning = value * 100d;
            }
        }

        /// <summary>
        /// If numbers continue vertically with vertical text or stay horizontal
        /// (If numbers are to be put in a single character block)
        /// Mostly relevant for east asian languages.
        /// </summary>
        public bool Kumimoji;

        /// <summary>
        /// Language to be used for UI controls.
        /// Overriden if alt-lang is set.
        /// </summary>
        internal string LanguageID;

        /// <summary>
        /// True specifies no spell or grammar check
        /// Default false.
        /// </summary>
        internal bool NoProof;

        /// <summary>
        /// Render-Only modification
        /// Wheter to normalize height when rendering. 
        /// Only changes visuals changes no actual values.
        /// False by default
        /// </summary>
        internal bool NormalizeH;

        /// <summary>
        /// Has this run been checked for smart tags
        /// </summary>
        internal bool SmartTagClean;

        /// <summary>
        /// Maps to some smart tag e.g. maps to a stock ticker symbol
        /// </summary>
        internal uint SmartTagID;

        /// <summary>
        /// Spacing between characters usually in 100s of a point
        /// Could technically also be a string including another unit of measurment see (ST_TextPoint)
        /// </summary>
        internal double Spacing;

        /// <summary>
        /// Font strike out type
        /// </summary>
        public eStrikeType Strike;

        //Minimum value is 100 or 1pt
        internal double RealSize = -1;

        /// <summary>
        /// Font size in points
        /// Maximum size 4000pts. Minimum 1pt
        /// will be clamped if set to larger
        /// </summary>
        public double Size
        {
            get
            {
                if (RealSize == -1) RealSize = 1;
                return RealSize / 100f;
            }
            set
            {
                var input = value;
                if (input < 1)
                {
                    input = 1;
                }
                else if (input > 4000)
                {
                    input = 4000;
                }
                RealSize = value * 100;
            }
        }

        /// <summary>
        /// Specifies wheter to underline the text 
        /// </summary>
        public eUnderLineType UnderLine;
    }
}
