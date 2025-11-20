using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables
{
    public abstract class FontTableBase
    {
        internal void Serialize(FontsBinaryWriter writer)
        {
            SerializeInternal(writer);
        }
        internal abstract void SerializeInternal(FontsBinaryWriter writer);

        internal abstract void Clear();

        public int GetLength()
        {
            using var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            SerializeInternal(writer);
            return (int)ms.Length;
        }
    }
}
