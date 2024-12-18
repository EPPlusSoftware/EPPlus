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
using OfficeOpenXml.Utils;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFontEx : LogFont
    {
        internal string FullName;
        internal string Style;
        internal string Script;

        internal LogFontEx() : base() {}

        internal LogFontEx(BinaryReader br) : base(br)
        {
            FullName = BinaryHelper.GetPotentiallyNullTerminatedString(br, 128, Encoding.Unicode);
            Style = BinaryHelper.GetPotentiallyNullTerminatedString(br, 64, Encoding.Unicode);
            Script = BinaryHelper.GetPotentiallyNullTerminatedString(br, 64, Encoding.Unicode);
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            BinaryHelper.WriteStringWithSetByteLength(bw, FullName, 128, Encoding.Unicode);
            BinaryHelper.WriteStringWithSetByteLength(bw, FullName, 64, Encoding.Unicode);
            BinaryHelper.WriteStringWithSetByteLength(bw, FullName, 64, Encoding.Unicode);
        }
    }
}
