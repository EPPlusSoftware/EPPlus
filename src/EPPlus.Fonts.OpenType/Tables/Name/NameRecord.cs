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
using EPPlus.Fonts.OpenType.FontLocalization;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    [DebuggerDisplay("{RecordType}: {Name}")]
    public class NameRecord : FontTableElement
    {
        /// <summary>
        /// Platform identifier code.
        /// </summary>
        public ushort platformId { get; set; }

        /// <summary>
        /// Platform-specific encoding identifier.
        /// </summary>
        public ushort encodingId { get; set; }

        /// <summary>
        /// Language identifier.
        /// </summary>
        public ushort languageID { get; set; }

        /// <summary>
        /// Name identifier.
        /// </summary>
        public ushort nameId { get; set; }

        public NameRecordTypes RecordType { get; set; }

        /// <summary>
        /// Name string length in bytes.
        /// </summary>
        public ushort length { get; set; }

        /// <summary>
        /// Name string offset in bytes from stringOffset.
        /// </summary>
        public ushort offset { get; set; }


        public string Name { get; set; }

        public LanguageMapping LanguageMapping { get; set; }

        public override string ToString()
        {
            if(!IsNullOrWhiteSpace(Name))
            {
                return Name;
            }
            else
            {
                return base.ToString();
            }
        }

        internal static bool IsNullOrWhiteSpace(string value)
        {
            return value == null || value.Trim().Length == 0;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(platformId);
            writer.WriteUInt16BigEndian(encodingId);
            writer.WriteUInt16BigEndian(languageID);
            writer.WriteUInt16BigEndian(nameId);
            writer.WriteUInt16BigEndian(length);
            writer.WriteUInt16BigEndian(offset);

        }
    }
}
