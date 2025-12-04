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

using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Utils.Platform;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    public static class OpenTypeFonts
    {
        private static object _syncRoot = new object();

        private static string GetWindowsFolder()
        {
            var winfolder = @"c:\Windows";
#if !NET35
            var wf = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(wf) && Directory.Exists(wf)) return wf;
#endif
            var ewf = Environment.GetEnvironmentVariable("WINDIR");
            if(!string.IsNullOrEmpty(ewf) && Directory.Exists(ewf))
            {
                winfolder = ewf;
            }
            return winfolder;
        }

        internal static List<string> winFontLocations = new List<string>()
        {
            Path.Combine(GetWindowsFolder(), "Fonts"),
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

        internal static List<string> GetLocationsCollection(IEnumerable<string> fontDirectories, bool searchSystemDirectories = true)
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

            return fontLocations;
        }

        static Dictionary<string, OpenTypeFont> CachedFonts;

        public static void ClearFontCache()
        {
            lock (_syncRoot)
            {
                if (CachedFonts != null && CachedFonts.Count > 0)
                {
                    CachedFonts.Clear();
                }
            }
        }


        public static OpenTypeFont GetFontDataOpen(IEnumerable<string> fontDirectories, string fontName, FontSubFamily subFamily = FontSubFamily.Regular, bool searchSystemDirectories = true)
        {
            lock (_syncRoot)
            {
                CachedFonts = CachedFonts == null ? new Dictionary<string, OpenTypeFont>() : CachedFonts;

                var fullName = fontName + "__" + subFamily;

                CachedFonts.TryGetValue(fullName, out OpenTypeFont cachedFont);
                if (cachedFont == null)
                {
                    //We do not have the font cached. Check what paths to search:
                    List<string> fontLocations = GetLocationsCollection(fontDirectories, searchSystemDirectories);

                    OpenTypeFont fontData = null;
                    foreach (var path in fontLocations)
                    {
                        var factory = new OpenTypeFontFactory(path);
                        fontData = factory.CreateBase(fontName, subFamily);
                        if (fontData != null)
                            break;
                    }

                    //another thread may have added it to the collection inbetween
                    if (!CachedFonts.ContainsKey(fullName))
                    {
                        CachedFonts.Add(fullName, fontData);
                    }
                    return fontData;
                }
                else
                {
                    return cachedFont;
                }
            }
        }

        public static List<OpenTypeFont> GetAllBaseFontData(List<string> fontDirectories, bool searchSystemDirectories = true, FontFormat? formatTarget = null)
        {
            var fontLocations = GetLocationsCollection(fontDirectories, searchSystemDirectories);

            List<OpenTypeFont> openTypeFontLst = new List<OpenTypeFont>();

            bool searchSpecificFormat = formatTarget != null;

            foreach (var path in fontLocations)
            {
                var scannedFonts = FontScanner.GetAllScannedFontsInPath(path);

                if(scannedFonts != null)
                {
                    foreach (var sf in scannedFonts)
                    {
                        var fontFactory = new OpenTypeFontFactory(sf.FilePath);

                        if (sf.SubFonts != null && sf.SubFonts.Any())
                        {
                            foreach (var subFont in sf.SubFonts)
                            {
                                if (searchSpecificFormat && subFont.Format != formatTarget)
                                {
                                    continue;
                                }

                                var familyName = string.IsNullOrEmpty(sf.FontFamilyName) ? subFont.FontFamilyName : sf.FontFamilyName;

                                var openFont = fontFactory.HandleScannedFontBase(subFont);

                                openTypeFontLst.Add(openFont);
                            }
                        }
                        else
                        {
                            if (searchSpecificFormat && sf.Format != formatTarget)
                            {
                                continue;
                            }

                            var openFont = fontFactory.HandleScannedFontBase(sf);
                            openTypeFontLst.Add(openFont);
                        }
                    }
                }
            }

            return openTypeFontLst;
        }

        public static OpenTypeFont GetFontData(IEnumerable<string> fontDirectories, string fontName, FontSubFamily subFamily, bool searchSystemDirectories = true)
        {
            lock (_syncRoot)
            {
                CachedFonts = CachedFonts == null ? new Dictionary<string, OpenTypeFont>() : CachedFonts;

            var fullName = fontName + "__" + subFamily.ToString();

                CachedFonts.TryGetValue(fullName, out OpenTypeFont cachedFont);
                if (cachedFont == null)
                {
                    //We do not have the font cached. Check what paths to search:
                    List<string> fontLocations = GetLocationsCollection(fontDirectories, searchSystemDirectories);

                    OpenTypeFont fontData = null;
                    foreach (var path in fontLocations)
                    {
                        var factory = new OpenTypeFontFactory(path);
                        fontData = factory.Create(fontName, subFamily);
                        if (fontData != null)
                            break;
                    }

                    //another thread may have added it to the collection inbetween
                    lock (_syncRoot)
                    {
                        if (!CachedFonts.ContainsKey(fullName))
                        {
                            CachedFonts.Add(fullName, fontData);
                        }
                    }


                    return fontData;
                }
                else
                {
                    return cachedFont;
                }
            }
        }
    }
}
