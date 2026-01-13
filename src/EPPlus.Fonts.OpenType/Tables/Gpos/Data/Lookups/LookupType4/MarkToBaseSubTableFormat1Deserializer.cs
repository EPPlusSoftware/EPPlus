/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Lookup Type 4 Deserializer
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// Deserializer for GPOS Lookup Type 4: MarkToBase Attachment Positioning
    /// </summary>
    internal class MarkToBaseSubTableFormat1Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public MarkToBaseSubTableFormat1Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        /// <summary>
        /// Deserializes a MarkToBase subtable from the current stream position.
        /// </summary>
        /// <param name="subtableStartOffset">Absolute offset where this subtable starts</param>
        /// <returns>Deserialized MarkToBaseSubTableFormat1</returns>
        public MarkToBaseSubTableFormat1 Deserialize(long subtableStartOffset)
        {
            _reader.BaseStream.Seek(subtableStartOffset, SeekOrigin.Begin);

            var subtable = new MarkToBaseSubTableFormat1
            {
                SubtableFormat = _reader.ReadUInt16BigEndian(),
                MarkCoverageOffset = _reader.ReadUInt16BigEndian(),
                BaseCoverageOffset = _reader.ReadUInt16BigEndian(),
                MarkClassCount = _reader.ReadUInt16BigEndian(),
                MarkArrayOffset = _reader.ReadUInt16BigEndian(),
                BaseArrayOffset = _reader.ReadUInt16BigEndian()
            };

            // Read MarkCoverage table
            if (subtable.MarkCoverageOffset > 0)
            {
                long coveragePos = subtableStartOffset + subtable.MarkCoverageOffset;
                subtable.MarkCoverage = ReadCoverage(coveragePos);
            }

            // Read BaseCoverage table
            if (subtable.BaseCoverageOffset > 0)
            {
                long coveragePos = subtableStartOffset + subtable.BaseCoverageOffset;
                subtable.BaseCoverage = ReadCoverage(coveragePos);
            }

            // Read MarkArray
            if (subtable.MarkArrayOffset > 0)
            {
                long markArrayPos = subtableStartOffset + subtable.MarkArrayOffset;
                subtable.MarkArray = ReadMarkArray(markArrayPos);
            }

            // Read BaseArray
            if (subtable.BaseArrayOffset > 0)
            {
                long baseArrayPos = subtableStartOffset + subtable.BaseArrayOffset;
                subtable.BaseArray = ReadBaseArray(baseArrayPos, subtable.MarkClassCount);
            }

            return subtable;
        }

        /// <summary>
        /// Reads a Coverage table (Format 1 or 2)
        /// </summary>
        private CoverageTable ReadCoverage(long coverageOffset)
        {
            _reader.BaseStream.Seek(coverageOffset, SeekOrigin.Begin);
            ushort coverageFormat = _reader.ReadUInt16BigEndian();

            if (coverageFormat == 1)
            {
                return new CoverageTableFormat1Deserializer(_reader).Deserialize(coverageOffset);
            }
            else if (coverageFormat == 2)
            {
                return new CoverageTableFormat2Deserializer(_reader).Deserialize(coverageOffset);
            }

            return null;
        }

        /// <summary>
        /// Reads a MarkArray table
        /// </summary>
        private MarkArray ReadMarkArray(long markArrayStart)
        {
            _reader.BaseStream.Seek(markArrayStart, SeekOrigin.Begin);

            var markArray = new MarkArray
            {
                MarkCount = _reader.ReadUInt16BigEndian()
            };

            markArray.Records = new MarkRecord[markArray.MarkCount];

            // Read MarkRecords
            for (int i = 0; i < markArray.MarkCount; i++)
            {
                markArray.Records[i] = new MarkRecord
                {
                    MarkClass = _reader.ReadUInt16BigEndian(),
                    MarkAnchorOffset = _reader.ReadUInt16BigEndian()
                };
            }

            // Read Anchor tables for each mark
            for (int i = 0; i < markArray.MarkCount; i++)
            {
                if (markArray.Records[i].MarkAnchorOffset > 0)
                {
                    long anchorPos = markArrayStart + markArray.Records[i].MarkAnchorOffset;
                    markArray.Records[i].MarkAnchor = ReadAnchor(anchorPos);
                }
            }

            return markArray;
        }

        /// <summary>
        /// Reads a BaseArray table
        /// </summary>
        private BaseArray ReadBaseArray(long baseArrayStart, ushort markClassCount)
        {
            _reader.BaseStream.Seek(baseArrayStart, SeekOrigin.Begin);

            var baseArray = new BaseArray
            {
                BaseCount = _reader.ReadUInt16BigEndian()
            };

            baseArray.Records = new BaseRecord[baseArray.BaseCount];

            // Read BaseRecords (offsets)
            for (int i = 0; i < baseArray.BaseCount; i++)
            {
                baseArray.Records[i] = new BaseRecord
                {
                    BaseAnchorOffsets = new ushort[markClassCount],
                    BaseAnchors = new AnchorTable[markClassCount]
                };

                // Read offsets for all mark classes
                for (int classIndex = 0; classIndex < markClassCount; classIndex++)
                {
                    baseArray.Records[i].BaseAnchorOffsets[classIndex] = _reader.ReadUInt16BigEndian();
                }
            }

            // Read Anchor tables
            for (int i = 0; i < baseArray.BaseCount; i++)
            {
                for (int classIndex = 0; classIndex < markClassCount; classIndex++)
                {
                    ushort offset = baseArray.Records[i].BaseAnchorOffsets[classIndex];
                    if (offset > 0)
                    {
                        long anchorPos = baseArrayStart + offset;
                        baseArray.Records[i].BaseAnchors[classIndex] = ReadAnchor(anchorPos);
                    }
                }
            }

            return baseArray;
        }

        /// <summary>
        /// Reads an Anchor table (Format 1, 2, or 3)
        /// </summary>
        private AnchorTable ReadAnchor(long anchorStart)
        {
            _reader.BaseStream.Seek(anchorStart, SeekOrigin.Begin);

            var anchor = new AnchorTable
            {
                AnchorFormat = _reader.ReadUInt16BigEndian(),
                XCoordinate = _reader.ReadInt16BigEndian(),
                YCoordinate = _reader.ReadInt16BigEndian()
            };

            if (anchor.AnchorFormat == 2)
            {
                // Format 2: includes anchor point index
                anchor.AnchorPoint = _reader.ReadUInt16BigEndian();
            }
            else if (anchor.AnchorFormat == 3)
            {
                // Format 3: includes device table offsets (we don't process them yet)
                anchor.XDeviceOffset = _reader.ReadUInt16BigEndian();
                anchor.YDeviceOffset = _reader.ReadUInt16BigEndian();
            }

            return anchor;
        }
    }
}