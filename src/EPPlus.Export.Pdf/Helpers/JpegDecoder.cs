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
    internal static class JpegDecoder
    {
        internal static bool IsJpeg(byte[] d) => d != null && d.Length > 2 && d[0] == 0xFF && d[1] == 0xD8;

        internal static void ReadJpegInfo(byte[] d, out int width, out int height, out int components, out bool adobe)
        {
            width = 0; height = 0; components = 3; adobe = false;
            if (d == null || d.Length < 4 || d[0] != 0xFF || d[1] != 0xD8) return;
            int i = 2;
            while (i + 1 < d.Length)
            {
                if (d[i] != 0xFF)
                {
                    i++;
                    continue;
                }
                byte marker = d[i + 1];
                if (marker == 0xFF)
                {
                    i++;
                    continue;
                }
                if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
                {
                    i += 2; continue;
                }
                if (i + 3 >= d.Length) return;
                int segLen = (d[i + 2] << 8) | d[i + 3];
                if (marker == 0xEE && i + 8 < d.Length &&
                    d[i + 4] == (byte)'A' && d[i + 5] == (byte)'d' && d[i + 6] == (byte)'o' &&
                    d[i + 7] == (byte)'b' && d[i + 8] == (byte)'e')
                {
                    adobe = true;
                }
                if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
                {
                    if (i + 9 >= d.Length) return;
                    height = (d[i + 5] << 8) | d[i + 6];
                    width = (d[i + 7] << 8) | d[i + 8];
                    components = d[i + 9];
                    return;
                }
                if (segLen < 2) return;
                i += 2 + segLen;
            }
        }
    }
}
