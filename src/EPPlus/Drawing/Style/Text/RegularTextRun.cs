/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System.Drawing;
using System.Reflection.Emit;
using System.Xml;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A richtext part
    /// </summary>
    internal class RegularTextRun
    {
        internal RegularTextRun()
        {
            TextCharacterAttributes Attributes = new TextCharacterAttributes();
        }

        internal TextCharacterAttributes Attributes;

        #region Attributes
        /// <summary>
        /// Bold text
        /// </summary>
        public bool Bold { get => Attributes.Bold; set => Attributes.Bold = value; }

        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public double Baseline { get => Attributes.Baseline; set => Attributes.Baseline = value; }

        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public eTextCapsType Capitalization { get => Attributes.Capitalization; set => Attributes.Capitalization = value; }

        /// <summary>
        /// Italic text
        /// </summary>
        public bool Italic { get => Attributes.Italic; set => Attributes.Italic = value; }

        /// <summary>
        /// The minimum font size at which character kerning occurs
        /// </summary>
        public double Kerning { get => Attributes.Kerning; set => Attributes.Kerning = value; }

        public double Spacing { get => Attributes.Spacing; set => Attributes.Spacing = value; }

        /// <summary>
        /// Strike-out text
        /// </summary>
        public eStrikeType Strike { get => Attributes.Strike; set => Attributes.Strike = value; }

        /// <summary>
        /// Fontsize
        /// Spans from 0-4000
        /// </summary>
        public double FontSize { get => Attributes.Size; set => Attributes.Size = value; }

        /// <summary>
        /// Underlined text
        /// </summary>
        public eUnderLineType UnderLine { get => Attributes.UnderLine; set => Attributes.UnderLine = value; }

        #endregion Attributes

        #region Properties

        #region LineProperties
        //TODO: Line Properties
        #endregion LineProperties

        #region Basic Fill
        ExcelDrawingFill _fill;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill;

        //Below is quick-access to the drawing fill

        /// <summary>
        /// Sets the default color of the text.
        /// This sets the Fill to a SolidFill with the specified color.
        /// <remark>
        /// Use the Fill property for more options
        /// </remark>
        /// </summary>
        public Color Color;
        #endregion Basic fill

        //TODO: EFFECTS

        //internal Color HighLight;

        public Color UnderLineColor;
        //TODO: UnderLineLineProperties

        #region FontNodes
        /// <summary>
        /// The latin typeface name
        /// </summary>
        public string LatinFont;

        /// <summary>
        /// The East Asian typeface name
        /// </summary>
        public string EastAsianFont;

        /// <summary>
        /// The complex font typeface name
        /// </summary>
        public string ComplexFont;

        /// <summary>
        /// The symbol font typeface name
        /// </summary>
        public string SymbolFont;

        #endregion FontNodes

        //TODO:
        #region HyperLink
        #endregion Hyperlink

        /// <summary>
        /// Right to left
        /// If ommitted it returns false AKA (left-to-right)
        /// </summary>
        internal bool rtl;

        //TODO:
        #region ExtLst-OfficeArtExtensionList
        #endregion ExtLst-OfficeArtExtensionList

        #endregion Properties

        /// <summary>
        /// Actual text for the text run
        /// </summary>
        internal string t;
    }
}
