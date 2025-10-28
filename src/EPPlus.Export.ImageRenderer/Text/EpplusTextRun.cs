/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace EPPlusImageRenderer.Text
{
    internal class EpplusTextRun
    {
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public eTextCapsType Capitalization;

        /// <summary>
        /// The minimum font size at which character kerning occurs
        /// </summary>
        public double Kerning;

        /// <summary>
        /// Fontsize
        /// Spans from 0-4000
        /// </summary>
        public double FontSize;

        /// <summary>
        /// The spacing between between characters
        /// </summary>
        public double Spacing;

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public double Baseline;

        /// <summary>
        /// FontBold text
        /// </summary>
        public bool Bold;

        /// <summary>
        /// FontItalic text
        /// </summary>
        public bool Italic;

        /// <summary>
        /// FontStrike-out text
        /// </summary>
        public eStrikeType Strike;

        /// <summary>
        /// Underlined text
        /// </summary>
        public eUnderLineType UnderLine;
    }
}
