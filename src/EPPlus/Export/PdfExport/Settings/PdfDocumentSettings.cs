using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
