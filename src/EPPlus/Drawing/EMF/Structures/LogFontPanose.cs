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
using OfficeOpenXml.Utils.FileUtils;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFontPanose : LogFont
    {
        string FullName;
        string Style;
        int Version;
        uint StyleSize;
        int Match;
        int Reserved;
        int VendorId;
        uint Culture;
        Panose PanoseObj;
        byte[] Padding;

        internal LogFontPanose (BinaryReader br) : base(br)
        {
            FullName = BinaryHelper.GetPotentiallyNullTerminatedString(br, 128, Encoding.Unicode);
            Style = BinaryHelper.GetPotentiallyNullTerminatedString(br, 64, Encoding.Unicode);
            Version = br.ReadInt32();
            StyleSize = br.ReadUInt32();
            Match = br.ReadInt32();
            Reserved = br.ReadInt32();
            VendorId = br.ReadInt32();
            Culture = br.ReadUInt32();
            PanoseObj = new Panose(br);
            Padding = br.ReadBytes(2);
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            BinaryHelper.WriteStringWithSetByteLength(bw, FullName, 128, Encoding.Unicode);
            BinaryHelper.WriteStringWithSetByteLength(bw, FullName, 64, Encoding.Unicode);
            bw.Write(Version);
            bw.Write(StyleSize);
            bw.Write(Match);
            bw.Write(Reserved);
            bw.Write(VendorId);
            bw.Write(Culture);
            PanoseObj.WriteBytes(bw);
            bw.Write(Padding);
        }
    }
}
