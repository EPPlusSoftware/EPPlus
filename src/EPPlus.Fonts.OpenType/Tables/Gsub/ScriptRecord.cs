using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class ScriptRecord : FontTableElement
    {
        public Tag ScriptTag { get; set; }
        public ushort ScriptOffset { get; set; }
        public ScriptTable ScriptTable { get; set; }

        // Implementerar den korrekta abstrakta metoden
        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Implementationen av serialisering kommer här.
            // TAG ScriptTag
            // USHORT ScriptOffset
            throw new NotImplementedException();
        }
    }
}
