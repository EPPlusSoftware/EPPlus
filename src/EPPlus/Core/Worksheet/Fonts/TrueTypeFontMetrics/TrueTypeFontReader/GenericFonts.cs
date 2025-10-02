using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.Utils.Platform;

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
    }
}
