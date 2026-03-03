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
using EPPlus.Fonts.OpenType.Utils.Platform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    internal static class DefaultFontLocations
    {
        private static string GetWindowsFolder()
        {
            var winfolder = @"c:\Windows";
#if !NET35
            var wf = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(wf) && Directory.Exists(wf)) return wf;
#endif
            var ewf = Environment.GetEnvironmentVariable("WINDIR");
            if (!string.IsNullOrEmpty(ewf) && Directory.Exists(ewf))
            {
                winfolder = ewf;
            }
            return winfolder;
        }

        internal static readonly List<string> winFontLocations = new List<string>()
        {
            Path.Combine(GetWindowsFolder(), "Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Windows\\Fonts"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\FontCache"),
        };

        internal static readonly List<string> linFontLocations = new List<string>()
        {
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            "/usr/share/X11/fonts",
            Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".fonts"),
            Path.Combine(Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", ".local"), "share"), "fonts"),
        };

        internal static readonly List<string> macFontLocations = new List<string>()
        {
            "/System/Library/Fonts",
            "/Library/Fonts",
            "/Network/Library/Fonts",
            Path.Combine(Path.Combine(Environment.GetEnvironmentVariable("HOME") ?? "~", "Library"), "Fonts"),
        };

        internal static List<string> GetLocationsCollection(IEnumerable<string> fontDirectories, bool searchSystemDirectories = true)
        {
            var fontLocations = new List<string>();
            fontLocations.AddRange(fontDirectories ?? Enumerable.Empty<string>());

            if (searchSystemDirectories)
            {
                var platform = PlatformUtils.GetPlatform();
                if (platform == PlatformUtils.OperatingSystem.Windows)
                    fontLocations.AddRange(winFontLocations);
                else if (platform == PlatformUtils.OperatingSystem.Mac)
                    fontLocations.AddRange(macFontLocations);
                else
                    fontLocations.AddRange(linFontLocations);
            }

            return fontLocations;
        }
    }
}
