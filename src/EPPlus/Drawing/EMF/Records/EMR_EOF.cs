/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EOF : EMR_RECORD
    {
        internal uint   nPalEntries;    //4
        internal uint   offPalEntries;  //4
        internal byte[] PaletteBuffer;  //Variable
        internal uint   SizeLast;       //4

        internal EMR_EOF(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            nPalEntries = br.ReadUInt32();
            offPalEntries = br.ReadUInt32();
            br.BaseStream.Position = position + offPalEntries;
            PaletteBuffer = br.ReadBytes((int)nPalEntries);
            SizeLast = br.ReadUInt32();
        }

        internal EMR_EOF()
        {
            Type = RECORD_TYPES.EMR_EOF;
            nPalEntries = 0;
            offPalEntries = 16;
            PaletteBuffer = new byte[nPalEntries];
            Size = (uint)(4 + 4 + 4 + 4 + 4 + PaletteBuffer.Length);
            SizeLast = Size;
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(nPalEntries);
            bw.Write(offPalEntries);
            bw.Write(PaletteBuffer);
            bw.Write(SizeLast);
        }
    }
}
