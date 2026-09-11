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
using EPPlus.Fonts.OpenType;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Settings
{
    internal class PdfDocumentSettings
    {
        internal OpenTypeFontEngine FontEngine;
        internal List<string> FontDirectories;
        internal bool SearchSystemDirectories;
        internal bool EmbeddFonts;
        internal string defaultFontName;
        internal int FirstPageNumber;
        internal bool Debug;
        internal bool PrintAsText;

        internal static PdfDocumentSettings From(PdfPageSettings s)
        {
            return new PdfDocumentSettings
            {
                FontEngine = s.FontEngine,
                FontDirectories = s.FontDirectories,
                SearchSystemDirectories = s.SearchSystemDirectories,
                EmbeddFonts = s.EmbeddFonts,
                defaultFontName = s.defaultFontName,
                FirstPageNumber = s.FirstPageNumber,
                Debug = s.Debug,
                PrintAsText = s.PrintAsText,
            };
        }

    }
}
