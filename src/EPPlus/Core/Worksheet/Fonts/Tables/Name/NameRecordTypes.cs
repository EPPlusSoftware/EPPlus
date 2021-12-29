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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Name
{
    /// <summary>
    /// The following name IDs are pre-defined, and they apply to all platforms unless indicated otherwise. 
    /// Name IDs 26 to 255, inclusive, are reserved for future standard names. Name IDs 256 to 32767, inclusive, 
    /// are reserved for font-specific names such as those referenced by a font’s layout features.
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/name#name-ids
    /// </summary>
    public enum NameRecordTypes : ushort
    {
        /// <summary>
        /// The copyright string from the font vendor
        /// </summary>
        CopyrightNotice = 0,
        /// <summary>
        /// The name the user sees. Times New Roman
        /// </summary>
        FontFamilyName = 1,
        /// <summary>
        /// The name of the style. Bold
        /// </summary>
        FontSubfamilyName = 2,
        /// <summary>
        /// A unique identifier that applications can store to identify the font being used. Monotype: Times New Roman Bold:1990
        /// </summary>
        UniqueFontIdentifier = 3,
        /// <summary>
        /// The complete, unique, human readable name of the font. This name is used by Windows. Times New Roman Bold
        /// </summary>
        FullFontName = 4,
        /// <summary>
        /// Release and version information from the font vendor. Version 1.00 June 1, 1990, initial release
        /// </summary>
        VersionString = 5,
        /// <summary>
        /// The name the font will be known by on a PostScript printer. TimesNewRoman-Bold
        /// </summary>
        PostScriptName = 6,
        /// <summary>
        ///  Trademark string. Times New Roman is a registered trademark of the Monotype Corporation.
        /// </summary>
        TradeMark = 7,
        /// <summary>
        /// Manufacturer. Monotype Corporation plc
        /// </summary>
        ManufacturerName = 8,
        /// <summary>
        /// Designer. Stanley Morison
        /// </summary>
        Designer = 9,
        /// <summary>
        /// Description. Designed in 1932 for the Times of London newspaper. Excellent readability and a narrow overall width, allowing more words per line than most fonts.
        /// </summary>
        Description = 10,
        /// <summary>
        /// URL of Vendor. http://www.monotype.com
        /// </summary>
        UrlVendor = 11,
        /// <summary>
        /// URL of Designer. http://www.monotype.com
        /// </summary>
        UrlDesigner = 12,
        /// <summary>
        ///  License Description. This font may be installed on all of your machines and printers, but you may not sell or give these fonts to anyone else.
        /// </summary>
        LicenseDescription = 13,
        /// <summary>
        /// License Info URL. http://www.monotype.com/license/
        /// </summary>
        LicenseInfoUrl = 14,
        Reserved1 = 15,
        /// <summary>
        /// Preferred Family. No name string present, since it is the same as name ID 1 (Font Family name).
        /// </summary>
        TypographicFamilyName = 16,
        /// <summary>
        /// Preferred Subfamily. No name string present, since it is the same as name ID 2 (Font Subfamily name).
        /// </summary>
        TypographicSubfamilyName = 17,
        /// <summary>
        /// Compatible Full (Macintosh only). No name string present, since it is the same as name ID 4 (Full name).
        /// </summary>
        CompatibleFull = 18,
        /// <summary>
        /// Sample text. The quick brown fox jumps over the lazy dog.
        /// </summary>
        SampleText = 19,
        /// <summary>
        /// PostScript CID findfont name. No name string present. Thus, the PostScript Name defined by name ID 6 should be used with the “findfont” invocation for locating the font in the context of a PostScript interpreter.
        /// </summary>
        PostScriptCID = 20,
        /// <summary>
        /// WWS family name: Since Times New Roman is a WWS font, this field does not need to be specified. If the font contained styles such as “caption”, “display”, “handwriting”, etc, that would be noted here.
        /// </summary>
        WWSfamilyName = 21,
        /// <summary>
        /// WWS subfamily name: Since Times New Roman is a WWS font, this field does not need to be specified.
        /// </summary>
        WWSsubfamilyName = 22,
        /// <summary>
        /// Light background palette name. No name string present, since this is not a color font.
        /// </summary>
        LightBackgroundPalette = 23,
        /// <summary>
        /// Dark background palette name. No name string present, since this is not a color font.
        /// </summary>
        DarkBackgroundPallet = 24,
        /// <summary>
        /// Variations PostScript name prefix. No name string present, since this is not a variable font.
        /// </summary>
        VariationsPostScriptName = 25
    }
}
