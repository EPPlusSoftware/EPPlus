/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/06/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Reads the raw font bytes for a given font face on disk.
    /// For TTF/OTF files this is just the file contents. For TTC collection files the
    /// individual font is extracted and returned as standalone TTF bytes.
    /// </summary>
    /// <remarks>
    /// This is a seam that lets DefaultFontResolver and other consumers be tested without
    /// touching the file system.
    /// </remarks>
    internal interface IFontFileReader
    {
        /// <summary>
        /// Reads the bytes for the given font face. For TTC files (OffsetInFile > 0) the
        /// returned bytes are a standalone TTF representation of the single font at that offset,
        /// not the raw TTC file contents.
        /// </summary>
        byte[] ReadFontBytes(FontFaceInfo face);
    }
}