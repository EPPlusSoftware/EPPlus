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
    internal class LogFont
    {
        //Read and written properties
        internal int Height;
        internal int Width;
        internal int Escapement;
        internal int Orientation;
        internal int Weight;
        internal byte Italic;
        internal byte Underline;
        internal byte StrikeOut;
        internal CharacterSet Set;
        internal byte OutPrecision;
        internal byte ClipPrecision;
        internal byte Quality;
        internal byte PitchAndFamily;
        internal string FaceName;

        //Simplified properties for viewing/editing
        internal FamilyFont fontFamily;
        internal Pitch pitchFont;

        private bool recalculateWidth = false;

        internal LogFont() { }

        internal LogFont(BinaryReader br)
        {
            Height = br.ReadInt32();
            Width = br.ReadInt32();
            if(Width == 0)
            {
                recalculateWidth = true;
            }
            Escapement = br.ReadInt32();
            Orientation = br.ReadInt32();
            Weight = br.ReadInt32();
            Italic = br.ReadByte();
            Underline = br.ReadByte();
            StrikeOut = br.ReadByte();
            Set = (CharacterSet)br.ReadByte();
            OutPrecision = br.ReadByte();
            ClipPrecision = br.ReadByte();
            Quality = br.ReadByte();
            PitchAndFamily = br.ReadByte();

            fontFamily = (FamilyFont)(PitchAndFamily >> 4);
            pitchFont = (Pitch)(PitchAndFamily & 0xF);

            //Should stop if encounters a terminating null
            FaceName = BinaryHelper.GetPotentiallyNullTerminatedString(br, 64, Encoding.Unicode);
        }

        internal virtual void WriteBytes(BinaryWriter bw)
        {
            bw.Write(Height);
            bw.Write(Width);
            bw.Write(Escapement);
            bw.Write(Orientation);
            bw.Write(Weight);
            bw.Write(Italic);
            bw.Write(Underline);
            bw.Write(StrikeOut);
            var byteTest = (byte)Set; 
            bw.Write(byteTest);
            bw.Write(OutPrecision);
            bw.Write(ClipPrecision);
            bw.Write(Quality);
            bw.Write(PitchAndFamily);
            BinaryHelper.WriteStringWithSetByteLength(bw, FaceName, 64, Encoding.Unicode);
        }
    }
}
