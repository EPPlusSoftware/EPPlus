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
    internal class DesignVector
    {
        internal uint Signature;
        internal uint NumAxes;
        internal uint[] Values;

        internal DesignVector() { }

        //Should only be used for a multiple master/OpenType font
        internal DesignVector(BinaryReader br)
        {
            Signature = br.ReadUInt32();
            if(Signature != 0x08007664)
            {
                throw new FileLoadException($"File corrupted! A DesignVector object MUST have signature {0x08007664}. Signature read was {Signature}");
            }
            NumAxes = br.ReadUInt32();

            Values = new uint[NumAxes];
            for(int i = 0; i < NumAxes; i++)
            {
                Values[i] = br.ReadUInt32();
            }
        }

        internal void WriteBytes(BinaryWriter bw)
        {
            bw.Write(Signature);
            bw.Write(NumAxes);
            if(NumAxes != 0)
            {
                foreach(var aValue in Values)
                {
                    bw.Write(aValue);
                }
            }
        }
    }
}
