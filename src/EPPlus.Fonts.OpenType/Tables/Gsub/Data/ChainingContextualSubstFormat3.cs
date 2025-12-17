using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    public class ChainingContextualSubstFormat3 : FontTableElement
    {
        // De tre bevakningsområdena
        public List<CoverageTable> BacktrackCoverages { get; set; } = new();
        public List<CoverageTable> InputCoverages { get; set; } = new();
        public List<CoverageTable> LookaheadCoverages { get; set; } = new();

        // Vad som ska hända vid matchning
        public List<SubstLookupRecord> SubstLookupRecords { get; set; } = new();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
