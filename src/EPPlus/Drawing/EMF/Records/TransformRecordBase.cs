using OfficeOpenXml.Drawing.EMF.Structures;
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class TransformRecordBase : EMR_RECORD
    {
        internal XForm xForm;
        //In most cases (MODIFYWORLDTRANSFORM type) this is ModifyWorldTransformMode.
        internal uint? TransformData = null;

        public TransformRecordBase(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            xForm = new XForm(br);
            if(br.BaseStream.Position - position < Size)
            {
                /* If MODIFYWORLDTRANSFORM
                1 = Identity
                2 = LeftMultiply,
                3 = RightMultiply,
                4 = Set(Set tranform To Data)
                */
                TransformData = br.ReadUInt32();
            }
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            xForm.WriteBytes(bw);
            if(TransformData != null)
            {
                bw.Write((uint)TransformData);
            }
        }
    }
}
