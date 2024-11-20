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
    internal class Panose
    {
        byte FamilyType;
        byte SerifStyle;
        byte Weight;
        byte Proportion;
        byte Contrast;
        byte StrokeVariation;
        byte ArmStyle;
        byte LetterForm;
        byte Midline;
        byte XHeight;

        internal Panose(BinaryReader br)
        {
            FamilyType = br.ReadByte();
            SerifStyle = br.ReadByte();
            Weight = br.ReadByte();
            Proportion = br.ReadByte();
            Contrast = br.ReadByte();
            StrokeVariation = br.ReadByte();
            ArmStyle = br.ReadByte();
            LetterForm = br.ReadByte();
            Midline = br.ReadByte();
            XHeight = br.ReadByte();
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(FamilyType);
            bw.Write(SerifStyle);
            bw.Write(Weight);
            bw.Write(Proportion);
            bw.Write(Contrast);
            bw.Write(StrokeVariation);
            bw.Write(ArmStyle);
            bw.Write(LetterForm);
            bw.Write(Midline);
            bw.Write(XHeight);
        }
    }
}
