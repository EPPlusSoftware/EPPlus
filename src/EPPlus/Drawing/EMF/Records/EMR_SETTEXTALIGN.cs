using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_SETTEXTALIGN : EMR_RECORD
    {
        internal TextAlignmentModeFlags TextAlignmentMode { get; set; }

        internal EMR_SETTEXTALIGN(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            TextAlignmentMode = (TextAlignmentModeFlags)br.ReadUInt32();
        }

        internal EMR_SETTEXTALIGN(TextAlignmentModeFlags textAlignmentMode)
        {
            TextAlignmentMode = textAlignmentMode;
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write((uint)TextAlignmentMode);
        }
    }
}
