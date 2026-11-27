/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2026         EPPlus Software AB           Raw pass-through table
 *************************************************************************************************/
using System;

namespace EPPlus.Fonts.OpenType.Tables
{
    /// <summary>
    /// A font table without a typed model. The bytes are kept and serialized verbatim.
    /// Used for tables that are copied unchanged into a subset (e.g. 'cvt ', 'fpgm', 'prep', 'gasp').
    /// </summary>
    internal class RawFontTable : FontTableBase
    {
        private readonly string _tag;
        private byte[] _data;

        internal RawFontTable(string tag, byte[] data)
        {
            if (tag == null) throw new ArgumentNullException(nameof(tag));
            if (tag.Length != 4) throw new ArgumentException("A table tag must be exactly four characters.", nameof(tag));
            if (data == null) throw new ArgumentNullException(nameof(data));

            _tag = tag;
            _data = data;
        }

        public override string Name => _tag;

        public override bool IsEssentialTable => false;

        /// <summary>
        /// The raw, unpadded table bytes.
        /// </summary>
        internal byte[] Data => _data;

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            writer.Write(_data);
        }

        internal override void Clear()
        {
            _data = new byte[0];
        }
    }
}