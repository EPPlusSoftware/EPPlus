using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Scanner
{
    internal static class FontScanner
    {
        internal static FontFormat? GetFormat(string file)
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

        internal static string[] TryGetFiles(string folder)
        {
            try
            {
                return Directory.GetFiles(folder);
            }
            catch (UnauthorizedAccessException) {}
            catch (DirectoryNotFoundException) {}
            catch (IOException) {}
            return new string[0];
        }

        internal static IScannedFont ScanFor(string fontDirectoryPath, string fontFamily, string subFamily)
        {
            var files = TryGetFiles(fontDirectoryPath).Where(x => Path.GetExtension(x).ToLower() == ".ttf" || Path.GetExtension(x).ToLower() == ".ttc" || Path.GetExtension(x).ToLower() == ".otf");
            if (!files.Any())
            {
                return default;
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
                            if (!string.IsNullOrEmpty(subFont.FontFamilyName) && subFont.FontFamilyName.ToLower() == fontFamily.ToLower())
                            {
                                if ((subFamily.ToLower() == "regular" || subFamily.ToLower() == "normal") && (sf.FontSubFamilyName.ToLower() == "regular" || sf.FontSubFamilyName.ToLower() == "normal"))
                                {
                                    return sf;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(sf.FontFamilyName) && sf.FontFamilyName.ToLower() == fontFamily.ToLower())
                        {
                            if ((subFamily.ToLower() == "regular" || subFamily.ToLower() == "normal") && (sf.FontSubFamilyName.ToLower() == "regular" || sf.FontSubFamilyName.ToLower() == "normal"))
                            {
                                return sf;
                            }
                        }
                    }
                }
            }
            return font;
        }
    }
}
