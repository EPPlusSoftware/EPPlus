using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables
{
    public abstract class FontTableElement
    {
        internal byte[] Serialize()
        {
            using var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            Serialize(writer);
            return ms.ToArray();
        }

        internal abstract void Serialize(FontsBinaryWriter writer);
    }
}
