using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Scanner;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.Utils.Platform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader
{
    internal static class GenericFonts
    {
        internal static List<string> winFontLocations = new List<string>()
        {
            Path.Combine(Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\Windows", "Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\FontCache"),
        };

        internal static List<string> linFontLocations = new List<string>()
        {
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            "/usr/share/X11/fonts",
            Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".fonts"),
            Path.Combine(Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".local"), "share"), "fonts"),
        };

        internal static List<string> macFontLocations = new List<string>()
        {
            "/System/Library/Fonts",
            "/Library/Fonts",
            "/Network/Library/Fonts",
            Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", "Library"), "Fonts"),
        };

        internal static TtfFont GetFontData(List<string> fontDirectories, bool searchSystemDirectories, string fontName, string subFamily = "Regular")
        {
            List<string> fontLocations = new List<string>();
            fontLocations.AddRange(fontDirectories);
            if (searchSystemDirectories)
            {
                var platform = PlatformUtils.GetPlatform();
                if (platform == PlatformUtils.OperatingSystem.Windows)
                {
                    fontLocations.AddRange(winFontLocations);
                }
                else if (platform == PlatformUtils.OperatingSystem.Mac)
                {
                    fontLocations.AddRange(macFontLocations);
                }
                else
                {
                    fontLocations.AddRange(linFontLocations);
                }
            }

            TtfFont fontData = null;
            foreach (var path in fontLocations)
            {
                var factory = new GenericFontFactory(path);
                fontData = factory.Create(fontName, subFamily);
                if (fontData != null)
                    break;
            }
            return fontData;
        }

        internal static List<TtfFont> GetAllFontData(List<string> fontDirectories, bool searchSystemDirectories = true)
        {
            List<string> fontLocations = new List<string>();
            fontLocations.AddRange(fontDirectories);
            if (searchSystemDirectories)
            {
                var platform = PlatformUtils.GetPlatform();
                if (platform == PlatformUtils.OperatingSystem.Windows)
                {
                    fontLocations.AddRange(winFontLocations);
                }
                else if (platform == PlatformUtils.OperatingSystem.Mac)
                {
                    fontLocations.AddRange(macFontLocations);
                }
                else
                {
                    fontLocations.AddRange(linFontLocations);
                }
            }

            List<TtfFont> trueTypeFontLst = new List<TtfFont>();

            TtfFont fontData = null;
            foreach (var path in fontLocations)
            {
                var files = FontScanner.TryGetFiles(path).Where(x => Path.GetExtension(x).ToLower() == ".ttf" || Path.GetExtension(x).ToLower() == ".ttc" || Path.GetExtension(x).ToLower() == ".otf");
                if (!files.Any())
                {
                    continue;
                }

                var factory = new GenericFontFactory(path);

                foreach (var file in files)
                {
                    using (var reader = new MyBinaryReader(File.OpenRead(file)))
                    {
                        var format = FontScanner.GetFormat(file);
                        if (!format.HasValue) continue;
                        var sf = new ScannedFont(reader, format.Value, file, 0);
                        sf.Format = format.Value;
                        reader.Close();

                        if (sf.SubFonts != null && sf.SubFonts.Any())
                        {
                            foreach (var subFont in sf.SubFonts)
                            {
                                var ttfFont = sf.TtcOffset.HasValue ?
                                    new TtfFont(new MyBinaryReader(File.OpenRead(sf.FilePath)), sf.TtcOffset.Value) :
                                    new TtfFont(new MyBinaryReader(File.OpenRead(sf.FilePath)));
                                trueTypeFontLst.Add(ttfFont);
                            }
                        }
                        else
                        {
                            var ttfFont = sf.TtcOffset.HasValue ?
                               new TtfFont(new MyBinaryReader(File.OpenRead(sf.FilePath)), sf.TtcOffset.Value) :
                               new TtfFont(new MyBinaryReader(File.OpenRead(sf.FilePath)));
                            trueTypeFontLst.Add(ttfFont);
                            //if (!string.IsNullOrEmpty(sf.FontFamilyName) && sf.FontFamilyName.ToLower() == fontFamily.ToLower())
                            //{
                            //    if ((subFamily.ToLower() == "regular" || subFamily.ToLower() == "normal") && (sf.FontSubFamilyName.ToLower() == "regular" || sf.FontSubFamilyName.ToLower() == "normal"))
                            //    {
                            //        return sf;
                            //    }
                            //}
                        }
                    }
                }

                //foreach (var fileName in files)
                //{
                //    var fontPath = fileName;
                //    var nameIndex = fontPath.LastIndexOf('\\') + 1;
                //    var fontNameOnly = fontPath.Substring(nameIndex, fontPath.LastIndexOf('.') - nameIndex);

                //    fontData = factory.Create(fontNameOnly, "Regular");
                //    if (fontData != null)
                //    {
                //        trueTypeFontLst.Add(fontData);
                //    }
                //}
            }

            return trueTypeFontLst;
        }

    }
}
