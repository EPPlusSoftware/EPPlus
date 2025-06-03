using FontLab1.Scanner;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FontLab1.GenericMeasurements
{
    internal class GenericFontFactory
    {
        public GenericFontFactory(string fontDirectoyPath)
        {
            _fontPath = fontDirectoyPath;
        }

        private List<IScannedFont> _cachedFonts;
        private readonly string _fontPath;

        private static float GetScaleFactor(string family, string subFamily)
        {
            if(family == "Tw Cen MT Condensed")
            {
                if (subFamily.StartsWith("Bold"))
                    return 1.18f;
                else if (subFamily == "Italic")
                    return 1.1f;
            }
            else if(subFamily.StartsWith("Bold"))
            {
                return 1.09f;
            }
            return 1.01f;
        }

        public TtfFont Create(string fontFamily, string subFamily)
        {
            var scannedFont = GetScannedFont(fontFamily, subFamily);
            if (scannedFont == null)
            {
                scannedFont = GetScannedFont(fontFamily, "Regular");
                if (scannedFont != null)
                {
                    var sf = GetScaleFactor(scannedFont.FontFamilyName, subFamily);
                    if (subFamily == "Bold" || subFamily == "Bold Italic")
                    {
                        return HandleScannedFont(scannedFont, subFamily, sf);
                    }
                    else if (subFamily == "Italic")
                    {
                        return HandleScannedFont(scannedFont, subFamily, sf);
                    }
                    return default(TtfFont);
                }
                return default(TtfFont);
            }
            else
            {
                if (scannedFont.SubFonts != null && scannedFont.SubFonts.Any())
                {
                    var subFont = scannedFont.SubFonts.FirstOrDefault(x => x.FontFamilyName == fontFamily && x.FontSubFamilyName == subFamily);
                    return HandleScannedFont(subFont, subFamily);
                }
                else
                {
                    return HandleScannedFont(scannedFont, subFamily);
                }
            }
        }

        private TtfFont HandleScannedFont(IScannedFont scannedFont, string subFamily, float widthScaleFactor = 1f)
        {
            return scannedFont.TtcOffset.HasValue ?
                new TtfFont(new MyBinaryReader(File.OpenRead(scannedFont.FilePath)), scannedFont.TtcOffset.Value) :
                new TtfFont(new MyBinaryReader(File.OpenRead(scannedFont.FilePath)));
        }

        private IScannedFont GetScannedFont(string family, string subFamily)
        {
            return FontScanner.ScanFor(_fontPath, family, subFamily);
        }
    }
}
