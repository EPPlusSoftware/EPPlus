/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS lookup handler interface
  09/07/2026         EPPlus Software AB           Support all wrapped lookup types with a handler,
                                                  and emit the inner lookup type instead of type 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Handlers
{
    /// <summary>
    /// Handler for GPOS Lookup Type 9: Extension Positioning.
    /// <para>
    /// The loader (<c>GposTableLoader.ReadExtensionPosSubTable</c>) already resolves the
    /// extension offset and stores the wrapped subtable directly in the lookup, while
    /// <see cref="LookupTable.LookupType"/> is left at 9. This handler therefore only has to
    /// determine the inner lookup type from the subtables and delegate to the handler for
    /// that type.
    /// </para>
    /// <para>
    /// The rewritten lookup is returned with the inner lookup type, not re-wrapped as type 9.
    /// Re-wrapping would require writing a real ExtensionPosFormat1 record (format, inner
    /// lookup type and a 32 bit offset) around every subtable. Since the subset is written
    /// from scratch and is small, there is no 16 bit offset overflow to avoid, so the
    /// wrapper serves no purpose.
    /// </para>
    /// </summary>
    internal class ExtensionPosHandler : IGposLookupHandler
    {
        public ushort LookupType => 9;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GposSubsetProcessor processor)
        {
            var unwrapped = Unwrap(lookup);
            if (unwrapped == null)
            {
                return;
            }

            var handler = processor.GetHandler(unwrapped.LookupType);
            if (handler != null)
            {
                handler.Discover(context, unwrapped, processor);
            }
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable lookup)
        {
            var unwrapped = Unwrap(lookup);
            if (unwrapped == null)
            {
                return null;
            }

            var handler = context.GposProcessor.GetHandler(unwrapped.LookupType);
            if (handler == null)
            {
                return null;
            }

            var rewritten = handler.Rewrite(context, unwrapped);
            if (rewritten == null || rewritten.SubTables.Count == 0)
            {
                return null;
            }

            // The inner handlers all copy LookupType from the lookup they are given, so the
            // rewritten lookup already carries the inner type. Return it as is.
            return rewritten;
        }

        /// <summary>
        /// Builds a lookup that represents the wrapped lookup, with the inner lookup type.
        /// Returns null when the lookup wraps a type that has no loaded subtables, which is the
        /// case for the positioning types the loader does not read yet (3, 5, 6, 7 and 8).
        /// </summary>
        private static LookupTable Unwrap(LookupTable lookup)
        {
            if (lookup.SubTables == null || lookup.SubTables.Count == 0)
            {
                return null;
            }

            ushort innerLookupType = 0;
            var innerSubTables = new List<FontTableElement>();

            foreach (var subTable in lookup.SubTables)
            {
                var subTableLookupType = GetLookupTypeOf(subTable);
                if (subTableLookupType == 0)
                {
                    continue;
                }

                if (innerLookupType == 0)
                {
                    innerLookupType = subTableLookupType;
                }
                else if (subTableLookupType != innerLookupType)
                {
                    // All subtables of a lookup must be of the same type per the OpenType
                    // specification. Ignore anything that does not match the first one rather
                    // than emitting a lookup whose type does not describe its subtables.
                    continue;
                }

                innerSubTables.Add(subTable);
            }

            if (innerLookupType == 0)
            {
                return null;
            }

            return new LookupTable
            {
                LookupType = innerLookupType,
                LookupFlag = lookup.LookupFlag,
                MarkFilteringSet = lookup.MarkFilteringSet,
                SubTables = innerSubTables
            };
        }

        /// <summary>
        /// Maps a loaded GPOS subtable to the lookup type it belongs to.
        /// Returns 0 for subtable types that cannot be mapped.
        /// </summary>
        private static ushort GetLookupTypeOf(FontTableElement subTable)
        {
            if (subTable is SinglePosSubTableFormat1 || subTable is SinglePosSubTableFormat2)
            {
                return 1;
            }

            if (subTable is PairPosSubTable)
            {
                // Covers both PairPosSubTableFormat1 and PairPosSubTableFormat2.
                return 2;
            }

            if (subTable is MarkToBaseSubTableFormat1)
            {
                return 4;
            }

            return 0;
        }
    }
}