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
    internal class LogFontExDv : LogFontEx
    {
        internal DesignVector dv;

        internal LogFontExDv() : base()
        {
            dv = new DesignVector();
        }
        internal LogFontExDv(BinaryReader br): base(br)
        {
            dv = new DesignVector(br);
        }
        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            dv.WriteBytes(bw);
        }
    }
}
