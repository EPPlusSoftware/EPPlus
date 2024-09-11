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
        internal byte[] offBmiSrc;
        internal uint   cbBmiSrc;
        internal byte[] offBitScr;
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
            offBmiSrc = br.ReadBytes(4);
            cbBmiSrc = br.ReadUInt32();
            offBitScr = br.ReadBytes(4);
            cbBitSrc = br.ReadUInt32();
            cxSrc = br.ReadBytes(4);
            cySrc = br.ReadBytes(4);

            BmiSrc = br.ReadBytes((int)cbBmiSrc);
            BitsSrc = br.ReadBytes((int)cbBitSrc);
        }
    }
}
