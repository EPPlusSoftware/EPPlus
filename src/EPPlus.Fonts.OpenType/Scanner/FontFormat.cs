/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Scanner
{

    /// <summary>
    /// Defines the supported font file formats that can be parsed or identified by the library.
    /// </summary>
    public enum FontFormat
    {
        /// <summary>
        /// TrueType Font (.ttf) — a widely used font format developed by Apple and Microsoft.
        /// </summa

        Ttf = 0,
        /// <summary>
        /// TrueType Collection (.ttc) — a container format that holds multiple TrueType fonts in a single file.
        /// </summary

        Ttc = 1,

        /// <summary>
        /// OpenType Font (.otf) — an extension of TrueType that supports advanced typographic features and PostScript outlines.
        /// </summary>
        Otf = 2
    }
}
