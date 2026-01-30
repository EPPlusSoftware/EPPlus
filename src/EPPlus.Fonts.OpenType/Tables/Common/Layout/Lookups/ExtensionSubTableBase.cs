using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Refactor: Common base for Extension
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups
{
    /// <summary>
    /// Base class for Extension Substitution (Type 7) and Extension Positioning (Type 9).
    /// Used to reference subtables that exceed the 16-bit offset limit.
    /// </summary>
    public abstract class ExtensionSubTableBase : FontTableElement
    {
        /// <summary>
        /// Gets or sets the lookup type of the subtable pointed to by the extension.
        /// For GSUB: 1-6 or 8. For GPOS: 1-8.
        /// </summary>
        public ushort ExtensionLookupType { get; set; }

        /// <summary>
        /// Gets or sets the actual subtable being extended.
        /// </summary>
        public FontTableElement ExtendedSubTable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long startPos = writer.BaseStream.Position;

            // 1. Write Format (always 1)
            writer.WriteUInt16BigEndian(1);

            // 2. Write the original Lookup Type
            writer.WriteUInt16BigEndian(ExtensionLookupType);

            // 3. Write Placeholder for 32-bit Offset (ULONG)
            long offsetPos = writer.BaseStream.Position;
            writer.WriteUInt32BigEndian(0);

            // 4. Serialize the extended subtable
            if (ExtendedSubTable != null)
            {
                long subTablePos = writer.BaseStream.Position;
                uint relativeOffset = (uint)(subTablePos - startPos);

                // Go back and write the 32-bit offset
                writer.BaseStream.Seek(offsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt32BigEndian(relativeOffset);

                // Return to end of subtable
                writer.BaseStream.Seek(subTablePos, System.IO.SeekOrigin.Begin);
                ExtendedSubTable.Serialize(writer);
            }
        }
    }
}