/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2026         EPPlus Software AB           OpenType font implementation
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Scanner
{
    internal static partial class FontScannerV2
    {
        /// <summary>
        /// Scans a TrueType Collection (.ttc) file and returns all contained font faces.
        /// Uses FontScannerCache to avoid re-scanning the same face multiple times.
        /// </summary>
        private static List<FontFaceInfo> ScanTtcFile(string filePath)
        {
            var faces = new List<FontFaceInfo>();

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (FontsBinaryReader reader = new FontsBinaryReader(fs))
            {
                try
                {
                    uint tag = reader.ReadUInt32BigEndian();
                    if (tag != 0x74746366) // "ttcf"
                    {
                        return faces; // Not a TTC file – return empty list
                    }

                    uint ttcVersion = reader.ReadUInt32BigEndian();
                    uint numFonts = reader.ReadUInt32BigEndian();

                    // Sanity check – prevent huge or corrupt TTC from causing issues
                    if (numFonts == 0 || numFonts > 1024)
                    {
                        return faces;
                    }

                    var offsets = new uint[numFonts];
                    for (int i = 0; i < numFonts; i++)
                    {
                        offsets[i] = reader.ReadUInt32BigEndian();
                    }

                    foreach (uint offset in offsets)
                    {
                        // Skip obviously invalid offsets
                        if (offset >= fs.Length)
                        {
                            continue;
                        }

                        try
                        {
                            var face = FontScannerCache.GetOrAdd(filePath, (long)offset, (path, off) =>
                            {
                                var f = FontScannerV2Core.ScanSingleFace(path, off);
                                // Explicitly mark as TTC – regardless of what ScanSingleFace detected
                                f.Format = FontFormat.Ttc;
                                return f;
                            });

                            faces.Add(face);
                        }
                        catch (Exception ex) when (
                            ex is EndOfStreamException ||
                            ex is IOException ||
                            ex is InvalidOperationException ||
                            ex is ArgumentException)
                        {
                            // These indicate a corrupt or malformed font face inside the TTC
                            System.Diagnostics.Debug.WriteLine(
                                $"[FontScannerV2] Failed to scan TTC face at offset 0x{offset:X8} in {filePath}\r\n" +
                                $"  Exception: {ex.GetType().Name}: {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            // Unexpected error – log fully
                            System.Diagnostics.Debug.WriteLine(
                                $"[FontScannerV2] UNEXPECTED ERROR scanning TTC face at offset 0x{offset:X8} in {filePath}\r\n" +
                                $"  Exception: {ex.GetType().Name}\r\n" +
                                $"  Message: {ex.Message}\r\n" +
                                $"  Stack: {ex.StackTrace}");
                        }
                    }
                }
                catch (Exception ex) when (ex is EndOfStreamException || ex is IOException)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[FontScannerV2] Failed to read TTC header in file: {filePath}\r\n" +
                        $"  Exception: {ex.GetType().Name}: {ex.Message}");
                }
            }



            return faces;
        }
    }
}
