/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Os2
{
    public class Os2Table
    {
        /// <summary>
        /// The version number for the OS/2 table: 0x0000 to 0x0005.
        /// The version number allows for identification of the precise contents and layout for the OS/2 table.
        /// </summary>
        public ushort version { get; set; }

        /// <summary>
        /// The Average Character Width parameter specifies the arithmetic average of the escapement (width) of all non-zero width glyphs in the font.
        /// </summary>
        public short xAvgCharWidth { get; set; }

        /// <summary>
        /// Indicates the visual weight (degree of blackness or thickness of strokes) of the characters in the font. Values from 1 to 1000 are valid.
        /// </summary>
        public ushort usWeightClass { get; set; }

        /// <summary>
        /// Indicates a relative change from the normal aspect ratio (width to height ratio) as specified by a font designer for the glyphs in a font.
        /// </summary>
        public ushort usWidthClass { get; set; }

        /// <summary>
        /// Indicates font embedding licensing rights for the font. The interpretation of flags is as follows:
        /// 0: Installable embedding: the font may be embedded, and may be permanently installed for use on a remote systems, or for use by other users.
        /// 2: Restricted License embedding: the font must not be modified, embedded or exchanged in any manner without first obtaining explicit permission of the legal owner.
        /// 4: Preview and Print embedding: the font may be embedded, and may be temporarily loaded on other systems for purposes of viewing or printing the document. Documents containing Preview and Print fonts must be opened “read-only”; no edits can be applied to the document.
        /// 8: Editable embedding: the font may be embedded, and may be temporarily loaded on other systems. As with Preview &amp; Print embedding, documents containing Editable fonts may be opened for reading. In addition, editing is permitted, including ability to format new text using the embedded font, and changes may be saved.
        /// </summary>
        public ushort fsType { get; set; }

        /// <summary>
        /// The recommended horizontal size in font design units for subscripts for this font.
        /// </summary>
        public short ySubscriptXSize { get; set; }

        /// <summary>
        /// The recommended vertical size in font design units for subscripts for this font.
        /// </summary>
        public short ySubscriptYSize { get; set; }

        /// <summary>
        /// The recommended horizontal offset in font design units for subscripts for this font.
        /// </summary>
        public short ySubscriptXOffset { get; set; }

        /// <summary>
        /// The recommended vertical offset in font design units from the baseline for subscripts for this font.
        /// </summary>
        public short ySubscriptYOffset { get; set; }

        /// <summary>
        /// The recommended horizontal size in font design units for superscripts for this font.
        /// </summary>
        public short ySuperscriptXSize { get; set; }

        /// <summary>
        /// The recommended vertical size in font design units for superscripts for this font.
        /// </summary>
        public short ySuperscriptYSize { get; set; }

        /// <summary>
        /// The recommended horizontal offset in font design units for superscripts for this font.
        /// </summary>
        public short ySuperscriptXOffset { get; set; }

        /// <summary>
        /// The recommended vertical offset in font design units from the baseline for superscripts for this font.
        /// </summary>
        public short ySuperscriptYOffset { get; set; }

        /// <summary>
        /// Thickness of the strikeout stroke in font design units.
        /// </summary>
        public short yStrikeoutSize { get; set; }

        /// <summary>
        /// The position of the top of the strikeout stroke relative to the baseline in font design units.
        /// </summary>
        public short yStrikeoutPosition { get; set; }

        /// <summary>
        /// This parameter is a classification of font-family design.
        /// The font class and font subclass are registered values assigned by IBM to each font family. This parameter is intended for use in selecting an alternate font when the requested font is not available. 
        /// </summary>
        public short sFamilyClass { get; set; }

        /// <summary>
        /// This 10-byte series of numbers is used to describe the visual characteristics of a given typeface.
        /// See https://docs.microsoft.com/en-us/typography/opentype/spec/os2#panose
        /// </summary>
        public short[] panose { get; set; }

        /// <summary>
        /// This field is used to specify the Unicode blocks or ranges encompassed by the font file 
        /// in 'cmap' subtables for platform 3, encoding ID 1 (Microsoft platform, Unicode BMP) and 
        /// platform 3, encoding ID 10 (Microsoft platform, Unicode full repertoire)
        /// See https://docs.microsoft.com/en-us/typography/opentype/spec/os2#ur
        /// </summary>
        public uint UnicodeRange1 { get; set; }

        public uint UnicodeRange2 { get; set; }

        public uint UnicodeRange3 { get; set; }

        public uint UnicodeRange4 { get; set; }

        /// <summary>
        /// The four-character identifier for the vendor of the given type face.
        /// </summary>
        public Tag archVendId { get; set; }

        /// <summary>
        /// Contains information concerning the nature of the font patterns
        /// See https://docs.microsoft.com/en-us/typography/opentype/spec/os2#fss
        /// </summary>
        public ushort fsSelection { get; set; }

        /// <summary>
        /// The minimum Unicode index (character code) in this font, according to the 'cmap' subtable for 
        /// platform ID 3 and platform- specific encoding ID 0 or 1. For most fonts supporting Win-ANSI 
        /// or other character sets, this value would be 0x0020. This field cannot represent supplementary 
        /// character values (codepoints greater than 0xFFFF). Fonts that support supplementary characters 
        /// should set the value in this field to 0xFFFF if the minimum index value is a supplementary character.
        /// </summary>
        public ushort usFirstCharIndex { get; set; }

        /// <summary>
        /// The maximum Unicode index (character code) in this font, according to the 'cmap' 
        /// subtable for platform ID 3 and encoding ID 0 or 1. This value depends on which character 
        /// sets the font supports. This field cannot represent supplementary character values (codepoints greater than 0xFFFF). 
        /// Fonts that support supplementary characters should set the value in this field to 0xFFFF.
        /// </summary>
        public ushort usLastCharIndex { get; set; }

        /// <summary>
        /// The typographic ascender for this font. This field should be combined with the sTypoDescender and sTypoLineGap values to determine default line spacing.
        /// See https://docs.microsoft.com/en-us/typography/opentype/spec/os2#stypoascender
        /// </summary>
        public short sTypoAscender { get; set; }

        /// <summary>
        /// The typographic descender for this font. This field should be combined with the sTypoAscender and sTypoLineGap values to determine default line spacing.
        /// https://docs.microsoft.com/en-us/typography/opentype/spec/os2#stypodescender
        /// </summary>
        public short sTypoDescender { get; set; }

        /// <summary>
        /// The typographic line gap for this font. This field should be combined with the sTypoAscender and sTypoDescender values to determine default line spacing.
        /// See https://docs.microsoft.com/en-us/typography/opentype/spec/os2#stypolinegap
        /// </summary>
        public short sTypoLineGap { get; set; }


    }
}
