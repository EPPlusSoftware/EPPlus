using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_INTERSECTCLIPRECT : EMR_RECORD
    {
        internal RectLObject Clip;

        internal EMR_INTERSECTCLIPRECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Clip = new RectLObject(br);
        }

        internal EMR_INTERSECTCLIPRECT()
        {
            Size = 24;
            Type = RECORD_TYPES.EMR_INTERSECTCLIPRECT;
            Clip = new RectLObject();
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            Clip.WriteBytes(bw);
        }
    }
}
