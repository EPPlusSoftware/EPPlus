using EPPlus.Fonts.OpenType.Subsetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal interface IGsubLookupHandler
    {
        ushort LookupType { get; }

        // Fas 1: Hitta vilka glyfer som påverkas
        void Discover(FontSubsettingContext context, LookupTable lookup);

        // Fas 2: Skapa en ny, filtrerad tabell
        LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup);
    }
}
