using System.Collections.Generic;

namespace FontLab1.Scanner
{
    internal interface IScannedFont
    {
        string FontFamilyName { get; }

        string FontSubFamilyName { get; set; }

        string FilePath { get; set; }

        FontFormat Format { get; set; }

        IEnumerable<ScannedFont> SubFonts { get; }

        long? TtcOffset { get; }
    }
}
