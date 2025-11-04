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
using System.IO;
using System;

namespace EPPlus.Fonts.OpenType.Tables.Kern
{
    public class KernSubTable : FontTableElement
    {
        public ushort version { get; set; }

        public ushort length { get; set; }

        public KernCoverage coverage { get; set; }

        public KernSubTableFormat0 Format0Subtable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Temp stream to calculate table length
            using var ms = new MemoryStream();
            using var tempWriter = new FontsBinaryWriter(ms);

            if (coverage.Format == 0 && Format0Subtable != null)
            {
                Format0Subtable.Serialize(tempWriter);
            }
            else
            {
                throw new NotSupportedException($"Unsupported kern subtable format: {coverage.Format}");
            }

            byte[] subtableData = ms.ToArray();
            length = (ushort)(subtableData.Length + 6); // 6 bytes for version, length, coverage

            // Write subtable-header
            writer.WriteUInt16BigEndian(version);
            writer.WriteUInt16BigEndian(length);
            writer.WriteUInt16BigEndian(coverage.RawValue);

            // Then write subtable-data
            writer.Write(subtableData);
        }
    }
}
