using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_STRETCHDIBITS : EMR_RECORD
    {
        internal byte[] Bounds;
        internal byte[] xDest;
        internal byte[] yDest;
        internal byte[] xSrc;
        internal byte[] ySrc;
        internal byte[] cxSrc;
        internal byte[] cySrc;
        internal uint   offBmiSrc;
        internal uint   cbBmiSrc;
        internal uint   offBitsSrc;
        internal uint   cbBitsSrc;
        internal byte[] UsageSrc;
        internal byte[] InternalBltRasterOperation;
        internal byte[] cxDest;
        internal byte[] cyDest;
        internal byte[] BmiSrc;
        internal byte[] BitsSrc;
        internal byte[] Padding;

        internal EMR_STRETCHDIBITS(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Bounds = br.ReadBytes(16);
            xDest = br.ReadBytes(4);
            yDest = br.ReadBytes(4);
            xSrc = br.ReadBytes(4);
            ySrc = br.ReadBytes(4);
            cxSrc = br.ReadBytes(4);
            cySrc = br.ReadBytes(4);
            offBmiSrc = br.ReadUInt32();
            cbBmiSrc = br.ReadUInt32();
            offBitsSrc = br.ReadUInt32();
            cbBitsSrc = br.ReadUInt32();
            UsageSrc = br.ReadBytes(4);
            InternalBltRasterOperation = br.ReadBytes(4);
            cxDest = br.ReadBytes(4);
            cyDest = br.ReadBytes(4);
            BmiSrc = br.ReadBytes((int)cbBmiSrc);
            BitsSrc = br.ReadBytes((int)cbBitsSrc);
            int padding = (int)((position + Size) - br.BaseStream.Position);
            if (padding < 0)
            {
                Padding = new byte[0];
                return;
            }
            Padding = br.ReadBytes(padding);
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(Bounds);
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
            bw.Write(Padding);
        }
    }
}
