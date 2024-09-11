using System.Drawing;
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EOF : EMR_RECORD
    {
        internal uint   nPalEntries;    //4
        internal uint   offPalEntries;  //4
        internal byte[] PaletteBuffer;  //Variable
        internal uint   SizeLast;       //4

        public EMR_EOF(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            nPalEntries = br.ReadUInt32();
            offPalEntries = br.ReadUInt32();
            br.BaseStream.Position = position + offPalEntries;
            PaletteBuffer = br.ReadBytes((int)nPalEntries);
            SizeLast = br.ReadUInt32();
        }

        public EMR_EOF()
        {
            Type = RECORD_TYPES.EMR_EOF;
            nPalEntries = 0;
            offPalEntries = 16;
            PaletteBuffer = new byte[nPalEntries];
            Size = (uint)(4 + 4 + 4 + 4 + 4 + PaletteBuffer.Length);
            SizeLast = Size;
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(nPalEntries);
            bw.Write(offPalEntries);
            bw.Write(PaletteBuffer);
            bw.Write(SizeLast);
        }
    }
}
