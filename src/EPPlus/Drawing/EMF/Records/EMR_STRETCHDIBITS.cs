using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_STRETCHDIBITS : EMR_RECORD
    {
        internal RectLObject Bounds;
        internal int xDest;
        internal int yDest;
        internal int xSrc;
        internal int ySrc;
        internal int cxSrc;
        internal int cySrc;
        internal uint   offBmiSrc;
        internal uint   cbBmiSrc;
        internal uint   offBitsSrc;
        internal uint   cbBitsSrc;
        internal uint UsageSrc;
        internal uint InternalBltRasterOperation;
        internal int cxDest;
        internal int cyDest;
        internal byte[] BmiSrc;
        internal byte[] BitsSrc;
        //internal byte[] Padding;

        internal EMR_STRETCHDIBITS(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            var startOfRecord = br.BaseStream.Position - 8;

            Bounds = new RectLObject(br);
            xDest = br.ReadInt32();
            yDest = br.ReadInt32();
            xSrc = br.ReadInt32();
            ySrc = br.ReadInt32();
            cxSrc = br.ReadInt32();
            cySrc = br.ReadInt32();
            offBmiSrc = br.ReadUInt32();
            cbBmiSrc = br.ReadUInt32();
            offBitsSrc = br.ReadUInt32();
            cbBitsSrc = br.ReadUInt32();
            UsageSrc = br.ReadUInt32();
            InternalBltRasterOperation = br.ReadUInt32();
            cxDest = br.ReadInt32();
            cyDest = br.ReadInt32();

            //There's undefined variable space here, ensure we reach the header
            var startOfHeader = startOfRecord + offBmiSrc;
            br.BaseStream.Position = startOfHeader;
            //BitmapHeader
            //BmiSrc = br.ReadBytes((int)cbBmiSrc);

            var bh = new BitmapHeader(br, cbBmiSrc);

            //There's undefined variable space here, ensure we reach the bitmapSpace
            var startOfBitmapBits = startOfRecord + offBitsSrc;
            br.BaseStream.Position = startOfBitmapBits;

            //Source bitmap bits
            BitsSrc = br.ReadBytes((int)cbBitsSrc);

            //int padding = (int)((position + Size) - br.BaseStream.Position);
            //if (padding < 0)
            //{
            //    Padding = new byte[0];
            //    return;
            //}
            //Padding = br.ReadBytes(padding);
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            Bounds.WriteBytes(bw);
            bw.Write(xDest);
            bw.Write(yDest);
            bw.Write(xSrc);
            bw.Write(ySrc);
            bw.Write(cxSrc);
            bw.Write(cySrc);
            bw.Write(offBmiSrc);
            bw.Write(cbBmiSrc);
            bw.Write(offBitsSrc);
            bw.Write(cbBitsSrc);
            bw.Write(UsageSrc); 
            bw.Write(InternalBltRasterOperation);
            bw.Write(cxDest);
            bw.Write(cyDest);
            bw.Write(BmiSrc);
            bw.Write(BitsSrc);
            //bw.Write(Padding);
        }
    }
}
