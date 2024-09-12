using System.IO;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EXTTEXTOUTW : EMR_RECORD
    {
        internal byte[] Bounds;
        internal byte[] iGraphicsMode;
        internal byte[] exScale;
        internal byte[] eyScale;
        internal byte[] Reference;
        internal uint   Chars;
        internal uint   offString;
        internal byte[] Options;
        internal byte[] Rectangle;
        internal uint   offDx;
        internal string StringBuffer;
        internal byte[] DxBuffer;

        private int padding = 0;

        public EMR_EXTTEXTOUTW(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Bounds = br.ReadBytes(16);
            iGraphicsMode = br.ReadBytes(4);
            exScale = br.ReadBytes(4);
            eyScale = br.ReadBytes(4);
            Reference = br.ReadBytes(8);
            Chars = br.ReadUInt32();
            offString = br.ReadUInt32();
            Options = br.ReadBytes(4);
            Rectangle = br.ReadBytes(16);
            offDx = br.ReadUInt32();
            br.BaseStream.Position = position + offString;
            StringBuffer = BinaryHelper.GetString(br, (Chars * 2), Encoding.Unicode);
            br.BaseStream.Position = position + offDx;
            DxBuffer = br.ReadBytes((int)(Size - offDx));
        }

        public EMR_EXTTEXTOUTW(string Text)
        {
            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
            Options = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            Rectangle = new byte[16] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF };
            offString = 4 + 4 + 16 + 4 + 4 + 4 + 8 + 4 + 4 + 4 + 16 + 4;
            StringBuffer = Text;
            Chars = (uint)StringBuffer.Length;
            offDx = offString + (uint)StringBuffer.Length * 2;
            if (offDx % 4 != 0)
            {
                padding = (int)offDx;
                offDx += 4 - (offDx % 4);
                padding = (int)(offDx) - padding;
            }
            DxBuffer = new byte[Text.Length * 4];
            int j = 0;
            for (int i=0; i<Text.Length; i++)
            {
                DxBuffer[j++] = 0x05;
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
            }
            Size = offDx + (uint)DxBuffer.Length;
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
            bw.Write(iGraphicsMode);
            bw.Write(exScale);
            bw.Write(eyScale);
            bw.Write(Reference);
            bw.Write(Chars);
            bw.Write(offString);
            bw.Write(Options);
            bw.Write(Rectangle);
            bw.Write(offDx);
            bw.Write(BinaryHelper.GetByteArray(StringBuffer, Encoding.Unicode));
            if (padding > 0)
            {
                bw.Write(new byte[padding]);
            }
            bw.Write(DxBuffer);
        }
    }
}
