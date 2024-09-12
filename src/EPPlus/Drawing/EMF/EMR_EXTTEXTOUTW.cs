using System;
using System.Collections.Generic;
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
            Reference = br.ReadBytes(8);        //Signed, koordinat för var texten börjar. 
            Chars = br.ReadUInt32();
            offString = br.ReadUInt32();
            Options = br.ReadBytes(4);
            Rectangle = br.ReadBytes(16);
            offDx = br.ReadUInt32();
            br.BaseStream.Position = position + offString;
            StringBuffer = BinaryHelper.GetString(br, (Chars * 2), Encoding.Unicode);
            br.BaseStream.Position = position + offDx;
            DxBuffer = br.ReadBytes((int)(Size - offDx));

            var changedSize = offDx - offString;
            changedSize -= (Chars * 2);
            if (changedSize > 0)
            {
                padding = (int)changedSize;
            }
        }

        byte[] GetDxBuffer()
        {
            DxBuffer = new byte[StringBuffer.Length * 4];
            int j = 0;
            for (int i = 0; i < StringBuffer.Length; i++)
            {
                DxBuffer[j++] = (byte)GetSpacingForChar(StringBuffer[i]);
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
            }
            return DxBuffer;
        }

        static int GetSpacingForChar(char aChar)
        {
            switch (aChar)
            {
                case 'a':
                    return 0x06;
                case 'b':
                    return 0x07;
                case 'c':
                    return 0x05;
                case 'd':
                    return 0x07;
                case 'e':
                    return 0x06;
                case 'f':
                    return 0x04;
                case 'g':
                    return 0x07;
                case 'h':
                    return 0x07;
                case 'i':
                    return 0x03;
                case 'j':
                    return 0x03;
                case 'k':
                    return 0x06;
                case 'l':
                    return 0x03;
                case 'm':
                    return 0x09;
                case 'n':
                    return 0x07;
                case 'o':
                    return 0x07;
                case 'p':
                    return 0x07;
                case 'q':
                    return 0x07;
                case 'r':
                    return 0x04;
                case 's':
                    return 0x05;
                case 't':
                    return 0x04;
                case 'u':
                    return 0x07;
                case 'v':
                    return 0x05;
                case 'w':
                    return 0x09;
                case 'x':
                    return 0x05;
                case 'y':
                    return 0x05;
                case 'z':
                    return 0x05;
                case 'A':
                    return 0x07;
                case 'B':
                    return 0x06;
                case 'C':
                    return 0x07;
                case 'D':
                    return 0x08;
                case 'E':
                    return 0x06;
                case 'F':
                    return 0x06;
                case 'G':
                    return 0x08;
                case 'H':
                    return 0x08;
                case 'I':
                    return 0x03;
                case 'J':
                    return 0x04;
                case 'K':
                    return 0x06;
                case 'L':
                    return 0x05;
                case 'M':
                    return 0x010;
                case 'N':
                    return 0x08;
                case 'O':
                    return 0x09;
                case 'P':
                    return 0x06;
                case 'Q':
                    return 0x08;
                case 'R':
                    return 0x07;
                case 'S':
                    return 0x06;
                case 'T':
                    return 0x06;
                case 'U':
                    return 0x08;
                case 'V':
                    return 0x07;
                case 'W':
                    return 0x011;
                case 'X':
                    return 0x06;
                case 'Y':
                    return 0x05;
                case 'Z':
                    return 0x06;
                default:
                    return 0x05;
            }
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

        public EMR_EXTTEXTOUTW(string Text, int x, int y)
        {
            Type = RECORD_TYPES.EMR_EXTTEXTOUTW;
            Bounds = new byte[16] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x4b, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x00 };
            iGraphicsMode = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            exScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            eyScale = new byte[4] { 0x00, 0x00, 0x00, 0x00 };

            var bx = BitConverter.GetBytes(x);
            var by = BitConverter.GetBytes(y);

            Reference = BinaryHelper.ConcatenateByteArrays(bx, by);

            //Reference = new byte[8] { 0x13, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00 };
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
            for (int i = 0; i < Text.Length; i++)
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
