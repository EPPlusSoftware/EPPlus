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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    internal class OpenTypeFontFactory
    {
        public OpenTypeFontFactory(string fontDirectoryPath)
        {
            _fontPath = fontDirectoryPath;
        }

        private readonly string _fontPath;

        //private static float GetScaleFactor(string family, string subFamily)
        //{
        //    if(family == "Tw Cen MT Condensed")
        //    {
        //        if (subFamily.StartsWith("Bold"))
        //            return 1.18f;
        //        else if (subFamily == "Italic")
        //            return 1.1f;
        //    }
        //    else if(subFamily.StartsWith("Bold"))
        //    {
        //        return 1.09f;
        //    }
        //    return 1.01f;
        //}

        public OpenTypeFont Create(string fontFamily, string subFamily)
        {
            var scannedFont = GetClosestScannedFont(fontFamily, subFamily);
            if (scannedFont != null)
            {
                return HandleScannedFont(scannedFont, subFamily);
            }
            else
            {
                return null;
            }
        }

        private OpenTypeFont HandleScannedFont(IScannedFont scannedFont, string subFamily, float widthScaleFactor = 1f)
        {
            var reader = new FontsBinaryReader(File.OpenRead(scannedFont.FilePath));
            return scannedFont.TtcOffset.HasValue ?
                new OpenTypeFont(reader, scannedFont.TtcOffset.Value, scannedFont.Format) :
                new OpenTypeFont(reader, scannedFont.Format);
        }

        public OpenTypeFont CreateBase(string fontFamily, string subFamily)
        {
            var scannedFont = GetClosestScannedFont(fontFamily, subFamily);
            if(scannedFont != null)
            {
                return HandleScannedFontBase(scannedFont, subFamily);
            }
            else
            {
                return null;
            }
        }

        internal OpenTypeFont HandleScannedFontBase(IScannedFont scannedFont, string subFamily, float widthScaleFactor = 1f)
        {
            var reader = new FontsBinaryReader(File.OpenRead(scannedFont.FilePath));

            return scannedFont.TtcOffset.HasValue ?
                new OpenTypeFont(reader, scannedFont.TtcOffset.Value, scannedFont.Format) :
                new OpenTypeFont(reader, scannedFont.Format);
        }

        //public List<OpenTypeFont> CreateAll()
        //{

        //}

        /// <summary>
        /// Get the specified font or at least the font family if only that can be found.
        /// </summary>
        /// <param name="family"></param>
        /// <param name="subFamily"></param>
        /// <returns></returns>
        private IScannedFont GetClosestScannedFont(string family, string subFamily)
        {
            return FontScanner.ScanForClosest(_fontPath, family, subFamily);
        }
    }
}
