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
using System;
using EPPlus.Fonts.OpenType.Subsetting;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents an Extension Substitution subtable (Lookup Type 7).
    /// This is used to reference subtables that exceed the 16-bit offset limit.
    /// </summary>
    public class ExtensionSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the lookup type of the subtable pointed to by the extension.
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

        /// <summary>
        /// Rewrites the extension by rewriting the inner subtable.
        /// </summary>
        internal ExtensionSubstSubTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            if (ExtendedSubTable == null) return null;

            // We delegate the rewrite to the specific type of the inner table
            FontTableElement rewrittenInner = null;

            if (ExtendedSubTable is SingleSubstSubTable single)
                rewrittenInner = single.Rewrite(context, oldLookup);
            else if (ExtendedSubTable is LigatureSubstSubTable ligature)
                rewrittenInner = ligature.Rewrite(context, oldLookup);
            else if (ExtendedSubTable is ChainingContextualSubstFormat3 contextual)
                rewrittenInner = contextual.Rewrite(context);
            // Add more types here as they are implemented

            if (rewrittenInner == null) return null;

            return new ExtensionSubstSubTable
            {
                ExtensionLookupType = this.ExtensionLookupType,
                ExtendedSubTable = rewrittenInner
            };
        }
    }
}