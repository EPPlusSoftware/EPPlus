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
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables
{
    public abstract class FontTableBase
    {
        internal byte[] Serialize(OpenTypeFont font)
        {
            return Serialize(font.GetSerializationContext());
        }

        internal byte[] Serialize(FontSerializationContext context)
        {
            using var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            SerializeInternal(writer, context);
            return ms.ToArray();
        }

        internal void Serialize(FontsBinaryWriter writer, FontSerializationContext context)
        {
            SerializeInternal(writer, context);
        }
        internal abstract void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context);

        internal abstract void Clear();

        public int GetLength(OpenTypeFont font)
        {
            return Serialize(font.GetSerializationContext()).Length;
        }

        public int GetLength(FontSerializationContext context)
        {
            return Serialize(context).Length;
        }

        public abstract string Name { get; }

        public abstract bool IsEssentialTable { get;  }
    }
}
