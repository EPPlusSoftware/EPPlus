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
namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a substitution lookup record within a contextual substitution table.
    /// It defines which lookup should be applied to a specific position in the input sequence.
    /// </summary>
    public class SubstLookupRecord
    {
        /// <summary>
        /// Gets or sets the zero-based index into the input glyph sequence where the substitution should be applied.
        /// </summary>
        public ushort SequenceIndex { get; set; }

        /// <summary>
        /// Gets or sets the index of the lookup in the GSUB LookupList that will be triggered for the specified sequence index.
        /// </summary>
        public ushort LookupListIndex { get; set; }
    }
}
