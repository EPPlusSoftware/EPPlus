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

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents an Extension Substitution subtable (Lookup Type 7).
    /// During the rewrite phase, these are typically unwrapped and stored as their 
    /// encapsulated lookup type (e.g., Type 4) unless the 16-bit offset limit is exceeded.
    /// </summary>
    public class ExtensionSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the lookup type of the subtable pointed to by the extension.
        /// </summary>
        public ushort ExtensionLookupType { get; set; }

        /// <summary>
        /// Gets or sets the actual substitution subtable encapsulated within the extension.
        /// </summary>
        public FontTableElement InnerSubTable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}