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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Scanner
{
    internal static class FontScanner
    {
        static Dictionary<string, ScannedFont> _scannedFontsCache;

        static Dictionary<string, ScannedFont> ScannedFontsCache
        {
            get
            {
                if (_scannedFontsCache == null)
                {
                    return new Dictionary<string, ScannedFont>();
                }
                else
                {
                    return _scannedFontsCache;
                }
            }
            set
            {
                _scannedFontsCache = value;
            }
        }

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
                if (Directory.Exists(folder))
                {
                    return Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
                }
            }
            catch (UnauthorizedAccessException) {}
            catch (DirectoryNotFoundException) {}
            catch (IOException) {}
            return new string[0];
        }

        internal static bool IsTargetFont(ScannedFont sf, string fontFamilyTarget, string subFamilyTarget)
        {
            if (sf.SubFonts != null && sf.SubFonts.Any())
            {
                foreach (var subFont in sf.SubFonts)
                {
                    if (!string.IsNullOrEmpty(subFont.FontFamilyName) && subFont.FontFamilyName.ToLower() == fontFamilyTarget.ToLower())
                    {
                        var subFamilyName = string.IsNullOrEmpty(sf.FontFamilyName) ? subFont.FontSubFamilyName : sf.FontSubFamilyName;

                        if (subFamilyTarget.ToLower() == subFamilyName.ToLower())
                        {
                            return true;
                        }
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(sf.FontFamilyName) && sf.FontFamilyName.ToLower() == fontFamilyTarget.ToLower())
                {
                    if (subFamilyTarget.ToLower() == sf.FontSubFamilyName.ToLower())
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        internal static IScannedFont ScanForClosest(string fontDirectoryPath, string fontFamily, string subFamily)
        {
            var scannedFont = ScanFor(fontDirectoryPath, fontFamily, subFamily);
            if (scannedFont == null)
            {
                scannedFont = ScanFor(fontDirectoryPath, fontFamily, "Regular");
                if (scannedFont != null)
                {
                    return scannedFont;
                }
                return default;
            }
            else
            {
                if (scannedFont.SubFonts != null && scannedFont.SubFonts.Any())
                {
                    var subFont = scannedFont.SubFonts.FirstOrDefault(x => x.FontFamilyName == fontFamily && x.FontSubFamilyName == subFamily);
                    return subFont;
                }
                else
                {
                    return scannedFont;
                }
            }
        }

        internal static IScannedFont ScanFor(string fontDirectoryPath, string fontFamily, string subFamily)
        {
            var font = default(IScannedFont);

            var fullName = fontFamily + "__" + subFamily;

            if (ScannedFontsCache.TryGetValue(fullName, out ScannedFont sf) == false)
            {
                var files = TryGetFiles(fontDirectoryPath).Where(x => Path.GetExtension(x).ToLower() == ".ttf" || Path.GetExtension(x).ToLower() == ".ttc" || Path.GetExtension(x).ToLower() == ".otf");

                if (!files.Any())
                {
                    return default(IScannedFont);
                }

                foreach (var file in files)
                {
                    using (var reader = new FontsBinaryReader(File.OpenRead(file)))
                    {
                        var format = GetFormat(file);
                        if (!format.HasValue) continue;

                        sf = new ScannedFont(reader, format.Value, file);
                        sf.Format = format.Value;

                        if(sf.FontFamilyName != null)
                        {
                            var individualFullName = sf.FontFamilyName;
                            if (sf.FontSubFamilyName != null)
                            {
                                individualFullName += "__" + sf.FontSubFamilyName;
                            }
                            if (ScannedFontsCache.ContainsKey(individualFullName) == false)
                            {
                                ScannedFontsCache.Add(individualFullName, sf);
                            }
                        }

                        if (IsTargetFont(sf, fontFamily, subFamily))
                        {
                            return sf;
                        }
                    }
                }
            }
            else
            {
                if (IsTargetFont(sf, fontFamily, subFamily))
                {
                    return sf;
                }
            }

            ////If we found the font-family but not the specific sub-family return the closest thing
            ////(likely 'regular' or 'normal' subfamily
            //if (ScannedFontsCache.TryGetValue(fontFamily, out ScannedFont sfBackup))
            //{
            //    return sfBackup;
            //}

            return font;
        }

        internal static List<ScannedFont> GetAllScannedFontsInPath(string fontDirectoryPath)
        {
            var files = TryGetFiles(fontDirectoryPath).Where(x => Path.GetExtension(x).ToLower() == ".ttf" || Path.GetExtension(x).ToLower() == ".ttc" || Path.GetExtension(x).ToLower() == ".otf");

            if (!files.Any())
            {
                return default;
            }

            List<ScannedFont> fontsInDirectory = new();

            foreach (var file in files)
            {
                using (var reader = new FontsBinaryReader(File.OpenRead(file)))
                {
                    var format = GetFormat(file);
                    if (!format.HasValue) continue;

                    var sf = new ScannedFont(reader, format.Value, file);
                    sf.Format = format.Value;

                    fontsInDirectory.Add(sf);
                }
            }

            return fontsInDirectory;
        }
    }
}
