/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
namespace EPPlus.Export.Pdf.Helpers
{
    /// <summary>
    /// Central guards for decoding untrusted embedded images. Dimensions come straight from
    /// attacker-controlled image headers, so every decoder must validate them (in long space,
    /// before allocating) to avoid int overflow, multi-GB allocations and decompression bombs.
    /// </summary>
    internal static class ImageLimits
    {
        /// <summary>Hard cap on either axis. Matches the existing BMP/GIF/TIFF/ICO caps.</summary>
        internal const int MaxDimension = 30000;

        /// <summary>Total pixel ceiling. 100 MP keeps width*height*4 well inside Int32 range.</summary>
        internal const long MaxPixels = 100_000_000;

        /// <summary>Upper bound on bytes produced by a single inflate (anti decompression-bomb).</summary>
        internal const long MaxDecodedBytes = 512L * 1024 * 1024;

        /// <summary>
        /// True only when width/height are positive, within range, and width*height*bytesPerPixel
        /// stays within the pixel and byte ceilings. All arithmetic is done in long, so this never
        /// overflows and can be called safely right before an allocation.
        /// </summary>
        internal static bool IsAllowed(int width, int height, int bytesPerPixel)
        {
            if (width <= 0 || height <= 0) return false;
            if (width > MaxDimension || height > MaxDimension) return false;
            long pixels = (long)width * height;
            if (pixels > MaxPixels) return false;
            if (pixels * bytesPerPixel > MaxDecodedBytes) return false;
            return true;
        }
    }
}