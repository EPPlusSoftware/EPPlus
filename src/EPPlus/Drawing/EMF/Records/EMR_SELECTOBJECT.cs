using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SELECTOBJECT : EMR_RECORD
    {
        internal uint ihObject;

        internal EMR_SELECTOBJECT(uint ihObject)
        {
            Type = RECORD_TYPES.EMR_SELECTOBJECT;
            Size = 12;
            this.ihObject = ihObject;
        }

        internal EMR_SELECTOBJECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            ihObject = br.ReadUInt32();
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihObject);
        }
    }
}
