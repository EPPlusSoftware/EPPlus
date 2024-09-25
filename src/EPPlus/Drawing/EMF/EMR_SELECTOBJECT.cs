using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SELECTOBJECT : EMR_RECORD
    {
        internal uint ihObject;

        public EMR_SELECTOBJECT(uint ihObject)
        {
            Type = RECORD_TYPES.EMR_SELECTOBJECT;
            Size = 12;
            this.ihObject = ihObject;
        }

        public EMR_SELECTOBJECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            ihObject = br.ReadUInt32();
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihObject);
        }
    }
}
