/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextLayoutEngine implementation
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System.Drawing;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Represents a text fragment with specific font properties.
    /// </summary>
    public class TextFragment
    {
        public string Text { get; set; }

        public MeasurementFont Font { get; set; }
        public ShapingOptions Options { get; set; }

        /// <summary>
        /// Store rich-text info.
        /// Nothing is supposed to be done with this within OpenType
        /// but we hold the data so users may more easily recognize what rich text this is in the output.
        /// </summary>
        public IRichTextInfoBase RichTextOptions { get; set; } = new RichTextDefaults();

        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }
    }

    /// <summary>
    /// Simple class to provide some kind of fallback/defaults
    /// </summary>
    public class RichTextDefaults : IRichTextInfoBase
    {
        internal RichTextDefaults()
        {
        }
        public bool IsItalic { get; set; } = false;

        public bool IsBold { get; set; } = false;

        public bool SubScript { get; set; } = false;

        public bool SuperScript { get; set; } = false;

        public int UnderlineType { get; set; } = -1;

        public int StrikeType { get; set; } = -1;

        public int Capitalization { get; set; } = -1;

        public Color UnderlineColor { get; set; }

        public Color FontColor { get; set; }

        //TODO Offset which is equal to 30% or -25% if Sub or Superscript are true?
    }
}
