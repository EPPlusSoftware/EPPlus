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

        public EMR_SELECTOBJECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            ihObject = br.ReadUInt32();
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihObject);
        }
    }
}
