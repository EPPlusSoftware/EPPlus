using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using System;
using System.IO;

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

            cxSrc = new byte[4] { 0x20, 0x00, 0x00, 0x00 };
            cySrc = new byte[4] { 0x20, 0x00, 0x00, 0x00 };

            BmiSrc = new byte[0]; //Empty because everything here is in BitsSrc

            Size = (uint)(offBmiSrc + BitsSrc.Length);
        }

        public void ChangeImage(byte[] bmp)
        {
            BitsSrc = new byte[bmp.Length - 14];
            Array.Copy(bmp, 14, BitsSrc, 0, bmp.Length-14);
            cbBitSrc = (uint)BitsSrc.Length;
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
            bw.Write(cbBmiSrc);
            bw.Write(offBmiSrc);
            bw.Write(cbBitSrc);
            bw.Write(cxSrc);
            bw.Write(cySrc);
            bw.Write(BmiSrc);
            bw.Write(BitsSrc);
        }
    }
}
