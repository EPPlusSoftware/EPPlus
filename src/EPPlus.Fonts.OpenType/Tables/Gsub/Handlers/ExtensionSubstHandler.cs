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
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal class ExtensionSubstHandler : IGsubLookupHandler
    {
        public ushort LookupType => 7;

        public void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor)
        {
            foreach (var subTable in lookup.SubTables.OfType<ExtensionSubstSubTable>())
            {
                // Extension-tabellen är bara en skal som pekar på den riktiga datan
                if (subTable.ExtendedSubTable == null) continue;

                // Vi skapar en "fejkad" LookupTable som vi kan skicka vidare till processorn
                // för att återanvända den logik vi redan har för t.ex. Typ 4.
                var dummyLookup = new LookupTable
                {
                    LookupType = subTable.ExtensionLookupType,
                    LookupFlag = lookup.LookupFlag,
                    SubTables = new List<FontTableElement> { subTable.ExtendedSubTable }
                };

                // Nu skickar vi tillbaka den inre tabellen till processorn
                processor.DiscoverLookup(context, dummyLookup);
            }
        }

        public LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            // För Rewrite låter vi din befintliga arkitektur sköta omskrivningen av subtabeller
            var newLookup = new LookupTable { LookupType = 7, LookupFlag = oldLookup.LookupFlag, SubTables = new List<FontTableElement>() };
            foreach (var subtable in oldLookup.SubTables.OfType<ExtensionSubstSubTable>())
            {
                var rewritten = subtable.Rewrite(context, oldLookup);
                if (rewritten != null) newLookup.SubTables.Add(rewritten);
            }
            return newLookup.SubTables.Count > 0 ? newLookup : null;
        }
    }
}
