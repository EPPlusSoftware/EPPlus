using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace FontLab1.Scanner
{
    internal static class FontScanner
    {
        private static FontFormat? GetFormat(string file)
        {
            var ext = Path.GetExtension(file).TrimStart('.').ToLowerInvariant();
            switch (ext)
            {
                case "ttc":
                    return FontFormat.Ttc;
                case "otf":
                    return FontFormat.Otf;
                case "ttf":
                    return FontFormat.Ttf;
                default:
                    return null;
            }
        }

        internal static IScannedFont ScanFor(string fontDirectoryPath, string fontFamily, string subFamily)
        {
            var files = Directory.GetFiles(fontDirectoryPath).Where(x => Path.GetExtension(x).ToLower() == ".ttf" || Path.GetExtension(x).ToLower() == ".ttc" || Path.GetExtension(x).ToLower() == ".otf");
            if (!files.Any())
            {
                return default(IScannedFont);
            }
            var font = default(IScannedFont);
            foreach (var file in files)
            {
                using (var reader = new MyBinaryReader(File.OpenRead(file)))
                {
                    var format = GetFormat(file);
                    if (!format.HasValue) continue;
                    var sf = new ScannedFont(reader, format.Value, file);
                    sf.Format = format.Value;
                    if(sf.SubFonts != null && sf.SubFonts.Any())
                    {
                        foreach(var subFont in sf.SubFonts)
                        {
                            if (!string.IsNullOrEmpty(subFont.FontFamilyName) && subFont.FontFamilyName.ToLower() == fontFamily.ToLower() && subFont.FontSubFamilyName.ToLower() == subFamily.ToLower())
                            {
                                return sf;
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(sf.FontFamilyName) && sf.FontFamilyName.ToLower() == fontFamily.ToLower() && sf.FontSubFamilyName.ToLower() == subFamily.ToLower())
                        {
                            return sf;
                        }
                    }
                }
            }
            return font;
        }
    }
}
