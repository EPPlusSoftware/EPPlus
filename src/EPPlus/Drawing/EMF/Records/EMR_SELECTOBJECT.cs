using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
