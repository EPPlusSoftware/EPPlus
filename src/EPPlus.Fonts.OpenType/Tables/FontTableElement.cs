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
using System.IO;

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

        internal void WriteRelativeOffset(FontsBinaryWriter writer, long startOfTable, long positionToUpdate)
        {
            long currentPos = writer.BaseStream.Position;

            // Beräkna offset (måste rymmas i en USHORT per OpenType spec)
            ushort relativeOffset = (ushort)(currentPos - startOfTable);

            // Gå tillbaka, skriv, och återställ position
            writer.BaseStream.Seek(positionToUpdate, SeekOrigin.Begin);
            writer.WriteUInt16BigEndian(relativeOffset);
            writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);
        }
    }
}
