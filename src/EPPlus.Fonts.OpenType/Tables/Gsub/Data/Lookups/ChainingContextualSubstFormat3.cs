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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Chaining Contextual Substitution Subtable Format 3 (Coverage-based).
    /// This format allows for substitutions based on the surrounding context of a glyph sequence.
    /// </summary>
    public class ChainingContextualSubstFormat3 : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of coverage tables for the backtrack sequence (glyphs appearing before the input).
        /// </summary>
        public List<CoverageTable> BacktrackCoverages { get; set; } = new();

        /// <summary>
        /// Gets or sets the list of coverage tables for the input sequence (glyphs to be substituted).
        /// </summary>
        public List<CoverageTable> InputCoverages { get; set; } = new();

        /// <summary>
        /// Gets or sets the list of coverage tables for the lookahead sequence (glyphs appearing after the input).
        /// </summary>
        public List<CoverageTable> LookaheadCoverages { get; set; } = new();

        /// <summary>
        /// Gets or sets the list of substitution lookup records that define which lookups to apply to the input sequence.
        /// </summary>
        public List<SubstLookupRecord> SubstLookupRecords { get; set; } = new();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}