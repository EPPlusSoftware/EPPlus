using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SETBKMODE : EMR_RECORD
    {
        internal uint BackgroundMode;

        public EMR_SETBKMODE(uint BackgroundMode)
        {
            Type = RECORD_TYPES.EMR_SETBKMODE;
            Size = 12;
            this.BackgroundMode = BackgroundMode;
        }

        public EMR_SETBKMODE(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            BackgroundMode = br.ReadUInt32();
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(BackgroundMode);
        }
    }
}
