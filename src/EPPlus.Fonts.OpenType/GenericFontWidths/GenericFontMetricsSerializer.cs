/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
  09/01/2026         EPPlus Software AB       Keep ranges instead of expanding them
 *************************************************************************************************/
using System;
using System.IO;
using System.Text;

namespace EPPlus.Fonts.OpenType.GenericFontWidths
{
    internal static class GenericFontMetricsSerializer
    {
        public static readonly Encoding FileEncoding = Encoding.UTF8;

        /// <summary>
        /// Highest .fmtr version this reader understands.
        ///
        /// The version field has been written since the format was introduced but never
        /// checked. Rejecting what we cannot read matters as soon as a version 2 exists,
        /// because the alternative is reading its extra fields as character data.
        /// </summary>
        internal const ushort MaxSupportedVersion = 2;

        public static SerializedFontMetrics Deserialize(Stream stream)
        {
            using (var reader = new BinaryReader(stream, FileEncoding))
            {
                var metrics = new SerializedFontMetrics();
                metrics.Version = reader.ReadUInt16();
                if (metrics.Version > MaxSupportedVersion)
                {
                    throw new InvalidDataException(
                        string.Format("Unsupported font metrics version {0}. This build reads up to version {1}.",
                                      metrics.Version, MaxSupportedVersion));
                }

                metrics.Family = (FontMetricsFamilies)reader.ReadUInt16();
                metrics.SubFamily = (FontSubFamilies)reader.ReadUInt16();
                metrics.LineHeight1em = reader.ReadSingle();
                metrics.DefaultWidthClass = (FontMetricsClass)reader.ReadByte();

                var nClassWidths = reader.ReadUInt16();
                if (nClassWidths == 0)
                {
                    ReadVerticalMetrics(reader, metrics);
                    metrics.Seal();
                    return metrics;
                }

                for (var x = 0; x < nClassWidths; x++)
                {
                    var cls = (FontMetricsClass)reader.ReadByte();
                    metrics.SetClassWidth(cls, reader.ReadSingle());
                }

                var nClasses = reader.ReadUInt16();
                for (var x = 0; x < nClasses; x++)
                {
                    var cls = (FontMetricsClass)reader.ReadByte();

                    var nRanges = reader.ReadUInt16();
                    for (var rngIx = 0; rngIx < nRanges; rngIx++)
                    {
                        var start = reader.ReadUInt16();
                        var end = reader.ReadUInt16();
                        // Kept as a range. The previous version expanded it here into one
                        // dictionary entry per character in the span.
                        metrics.AddRange(start, end, cls);
                    }

                    var nCharactersInClass = reader.ReadUInt16();
                    for (var y = 0; y < nCharactersInClass; y++)
                    {
                        metrics.AddCharacter(reader.ReadUInt16(), cls);
                    }
                }

                ReadVerticalMetrics(reader, metrics);
                metrics.Seal();
                return metrics;
            }
        }

        /// <summary>
        /// Reads the version 2 vertical metrics, which sit at the end of the file after the
        /// class section. Placing them last keeps the version 1 part of a version 2 file byte
        /// identical to a version 1 file.
        ///
        /// The line height is derived from the three rather than read from its own field. All
        /// three go through the generator's per-value adjustment, so a separately written total
        /// would no longer equal their sum and the baseline could end up outside the line box.
        /// </summary>
        private static void ReadVerticalMetrics(BinaryReader reader, SerializedFontMetrics metrics)
        {
            if (metrics.Version < 2)
            {
                metrics.ApproximateVerticalMetricsFromLineHeight();
                return;
            }

            metrics.Ascender1em = reader.ReadSingle();
            metrics.Descender1em = reader.ReadSingle();
            metrics.LineGap1em = reader.ReadSingle();
            metrics.LineHeight1em = metrics.Ascender1em + metrics.Descender1em + metrics.LineGap1em;
        }
    }
}