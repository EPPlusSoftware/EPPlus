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
        internal string stringBuffer;
        internal byte[] DxBuffer;

        private int padding = 0;

        internal string Text
        {
            get 
            {
                return stringBuffer;
            }
            set
            {
                stringBuffer = value;
                CalculateOffsets();
            }
        }

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
            //Encoding is specifically UTF-16LE meaning no BOM allowed and little endian
            stringBuffer = BinaryHelper.GetString(br, (Chars * 2), Encoding.Unicode);
            br.BaseStream.Position = position + offDx;
            DxBuffer = br.ReadBytes((int)(Size - offDx));

            var changedSize = offDx - offString;
            changedSize -= (Chars * 2);
            if (changedSize > 0)
            {
                padding = (int)changedSize;
            }
        }

        byte[] CalculateDxSpacing()
        {
            int j = 0;
            for (int i = 0; i < stringBuffer.Length; i++)
            {
                DxBuffer[j++] = (byte)GetSpacingForChar(stringBuffer[i]);
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
                DxBuffer[j++] = 0x00;
            }
            return DxBuffer;
        }

        internal static int GetSpacingForChar(char aChar)
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
                    return 0x0A;
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
                    return 0x0B;
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
            stringBuffer = Text;
            CalculateOffsets();
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
            stringBuffer = Text;
            CalculateOffsets();
        }

        private void CalculateOffsets()
        {
            Chars = (uint)stringBuffer.Length;
            offDx = offString + (uint)stringBuffer.Length * 2;

            padding = (int)offDx;
            offDx += 4 - (offDx % 4);
            padding = (int)(offDx) - padding;

            DxBuffer = new byte[stringBuffer.Length * 4];
            CalculateDxSpacing();
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
            bw.Write(BinaryHelper.GetByteArray(stringBuffer, Encoding.Unicode));
            if (padding > 0)
            {
                bw.Write(new byte[padding]);
            }
            bw.Write(DxBuffer);
        }


        //we can fit 32 of the smalest character, l   len = 96, but we go len 90, 30 char, range 0-47
        //We can fit 11 of the widest character, O.   len = 99 but we go len 90, 10 char, range 0-44

        //Create new class that takes a string, gets it's len
        //if len is bigger than 90, cut at 90
        //repeat until string is end or more than 3 rows.
        //For each string create new textbox which is a record collection cosisting of:
            /*
            EMF_EXTCREATEFONTINDIRECTW Font;
            EMF_SELECTOBJECT sel1;
            EMF_SELECTOBJECT bkmode;
            EMF_EXTTEXTOUTW text;
            EMF_SELECTOBJECT sel2;
            EMF_DELETEOBJECT del;
            */
        //Calculate ReferenceX in Text based on len. lower len higher value.
        //RefenceY is increased by 12.
        //Remove current textRecords
        //Add our Text records
        //Done
        }
}
