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
using System.IO;

namespace EPPlus.Fonts.OpenType
{
    internal static class OpenTypeFontFactory
    {
        public static OpenTypeFont CreateFromFace(FontFaceInfo face)
        {
            var stream = new FileStream(face.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var reader = new FontsBinaryReader(stream);
            reader.BaseStream.Position = face.OffsetInFile;

            var format = face.OffsetInFile > 0
               ? FontFormat.Ttf 
               : Path.GetExtension(face.FilePath).ToLowerInvariant() == ".otf"
                   ? FontFormat.Otf
                   : FontFormat.Ttf;

            return new OpenTypeFont(reader, format);
        }
    }
}