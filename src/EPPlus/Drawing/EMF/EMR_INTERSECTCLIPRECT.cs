using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_INTERSECTCLIPRECT : EMR_RECORD
    {
        internal RectLObject Clip;

        public EMR_INTERSECTCLIPRECT(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            Clip = new RectLObject(br);
        }

        public EMR_INTERSECTCLIPRECT()
        {
            Type = RECORD_TYPES.EMR_INTERSECTCLIPRECT;
            Clip = new RectLObject();
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            Clip.WriteBytes(bw);
        }
    }
}
