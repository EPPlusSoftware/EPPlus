/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
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
