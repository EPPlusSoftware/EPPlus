using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
