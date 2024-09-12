using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SETTEXTALIGN : EMR_RECORD
    {
        internal TextAlignmentModeFlags TextAlignmentMode { get; set; }

        public EMR_SETTEXTALIGN(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            TextAlignmentMode = (TextAlignmentModeFlags)br.ReadUInt32();
        }

        public EMR_SETTEXTALIGN(TextAlignmentModeFlags textAlignmentMode)
        {
            TextAlignmentMode = textAlignmentMode;
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write((uint)TextAlignmentMode);
        }
    }
}
