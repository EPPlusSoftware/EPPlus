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
    internal class EMR_SELECTOBJECT : EMR_RECORD
    {
        /// <summary>
        /// Index of a graphics object either in the EMF object table or stock object 
        /// </summary>
        internal uint ihObject;

        internal EMR_SELECTOBJECT(uint ihObject)
        {
            Type = RECORD_TYPES.EMR_SELECTOBJECT;
            Size = 12;
            this.ihObject = ihObject;
        }

        public EMR_SELECTOBJECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            ihObject = br.ReadUInt32();
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihObject);
        }
    }
}
