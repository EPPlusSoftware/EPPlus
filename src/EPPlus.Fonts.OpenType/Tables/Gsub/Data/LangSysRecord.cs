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
namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    /// <summary>
    /// Represents a Language System Record, which associates a specific language tag 
    /// with a Language System Table.
    /// </summary>
    public class LangSysRecord
    {
        /// <summary>
        /// Gets or sets the 4-byte Language System tag (e.g., 'SVE ' for Swedish), stored as a uint.
        /// </summary>
        public uint LangSysTag { get; set; }

        /// <summary>
        /// Gets or sets the actual Language System Table associated with this record.
        /// </summary>
        public LangSysTable LangSysTable { get; set; }
    }
}