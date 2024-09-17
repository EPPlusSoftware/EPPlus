using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using System;
using System.IO;
using static OfficeOpenXml.Drawing.OleObject.OleObjectDataStreams;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_STRETCHBLT : EMR_RECORD
    {
        internal byte[] Bounds;
        internal byte[] xDest;
        internal byte[] yDest;
        internal byte[] cxDest;
        internal byte[] cyDest;
        internal byte[] BitBltRasterOperation;
        internal byte[] xSrc;
        internal byte[] ySrc;
        internal byte[] XformSrc;
        internal byte[] BkColorSrc;
        internal byte[] UsageSrc;
        internal uint   offBmiSrc;
        internal uint   cbBmiSrc;
        internal uint   offBitScr;
        internal uint   cbBitSrc;
        internal byte[] cxSrc;
        internal byte[] cySrc;
        internal byte[] BmiSrc;
        internal byte[] BitsSrc;
        internal byte[] Padding;

        public EMR_STRETCHBLT(BinaryReader br, uint TypeValue) : base(br , TypeValue)
        {
            Bounds = br.ReadBytes(16);
            xDest = br.ReadBytes(4);
            yDest = br.ReadBytes(4);
            cxDest = br.ReadBytes(4);
            cyDest = br.ReadBytes(4);
            BitBltRasterOperation = br.ReadBytes(4);
            xSrc = br.ReadBytes(4);
            ySrc = br.ReadBytes(4);
            XformSrc = br.ReadBytes(24);
            BkColorSrc = br.ReadBytes(4);
            UsageSrc = br.ReadBytes(4);
            offBmiSrc = br.ReadUInt32();
            cbBmiSrc = br.ReadUInt32();
            offBitScr = br.ReadUInt32();
            cbBitSrc = br.ReadUInt32();
            cxSrc = br.ReadBytes(4);
            cySrc = br.ReadBytes(4);

            BmiSrc = br.ReadBytes((int)cbBmiSrc);
            BitsSrc = br.ReadBytes((int)cbBitSrc);

            int padding = (int)((position + Size) - br.BaseStream.Position);
            if (padding < 0)
            {
                Padding = new byte[0];
                return;
            }
            Padding = br.ReadBytes(padding);
        }

        public EMR_STRETCHBLT(byte[] bmp)
        {
            Type = RECORD_TYPES.EMR_STRETCHBLT;
            Bounds = new byte[16] { 0x20, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x3F, 0x00, 0x00, 0x00, 0x21, 0x00, 0x00, 0x00 };
            xDest = new byte[4] { 0x20, 0x00, 0x00, 0x00 };
            yDest = new byte[4] { 0x02, 0x00, 0x00, 0x00 };
            cxDest = new byte[4] { 0x20, 0x00, 0x00, 0x00 };
            cyDest = new byte[4] { 0x20, 0x00, 0x00, 0x00 };
            BitBltRasterOperation = new byte[4] { 0x46, 0x00, 0x66, 0x00 };
            xSrc = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            ySrc = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            XformSrc = new byte[24] { 0x00, 0x00, 0x80, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
            BkColorSrc = new byte[4] { 0xFF, 0xFF, 0xF, 0x00 };
            UsageSrc = new byte[4] { 0x00, 0x00, 0x00, 0x00 };
            offBmiSrc = 4 + 4 + 16 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 24 + 4 + 4 + 4 + 4 + 4 + 4 + 4 + 4;
            cbBmiSrc = 40;
            offBitScr = offBmiSrc + cbBmiSrc;

            ChangeImage(bmp);
        }

        public void ChangeImage(byte[] bmp)
        {
            byte[] bmpHeader = new byte[14];
            Array.Copy(bmp, 0, bmpHeader, 0, 14);

            byte[] bmpDIBHeaderSize = new byte[4];
            Array.Copy(bmp, 14, bmpDIBHeaderSize, 0, 4);

            int DIBHeaderSize = BitConverter.ToInt32(bmpDIBHeaderSize, 0);

            byte[] bmpDIBHeader = new byte[DIBHeaderSize];
            Array.Copy(bmp, 14, bmpDIBHeader, 0, DIBHeaderSize);


            //Get width and height from bmp image. This will make sure we display the full image.
            Array.Copy(bmpDIBHeader, 4, cxSrc, 0, 4);
            Array.Copy(bmpDIBHeader, 8, cySrc, 0, 4);

            int headerSize = DIBHeaderSize + 14;

            byte[] bmpPixelData = new byte[bmp.Length - headerSize];
            Array.Copy(bmp, headerSize, bmpPixelData, 0, bmp.Length - headerSize);

            BmiSrc = bmpDIBHeader;
            var headerSizeDiff = cbBmiSrc - bmpDIBHeader.Length;
            cbBmiSrc = (uint)bmpDIBHeader.Length;

            offBitScr = offBitScr - (uint)headerSizeDiff;

            cbBitSrc = (uint)bmpPixelData.Length;
            BitsSrc = bmpPixelData;

            Size = (uint)(offBitScr + cbBitSrc);
            int paddingBytes = (int)(4 - (Size % 4)) % 4;
            Padding = new byte[paddingBytes];

        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
            bw.Write(xDest);
            bw.Write(yDest);
            bw.Write(cxDest);
            bw.Write(cyDest);
            bw.Write(BitBltRasterOperation);
            bw.Write(xSrc);
            bw.Write(ySrc);
            bw.Write(XformSrc);
            bw.Write(BkColorSrc);
            bw.Write(UsageSrc);
            bw.Write(offBmiSrc);
            bw.Write(cbBmiSrc);
            bw.Write(offBitScr);
            bw.Write(cbBitSrc);
            bw.Write(cxSrc);
            bw.Write(cySrc);
            bw.Write(BmiSrc);
            bw.Write(BitsSrc);
            bw.Write(Padding);
        }
    }
}
