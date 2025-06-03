using FontLab1.FontLocalization;
using System.Diagnostics;

namespace FontLab1.Tables.Name
{
    [DebuggerDisplay("{RecordType}: {Name}")]
    internal class NameRecord
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
    }
}
