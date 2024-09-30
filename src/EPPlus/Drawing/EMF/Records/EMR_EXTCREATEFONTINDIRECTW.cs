using System;
using System.IO;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EXTCREATEFONTINDIRECTW : EMR_RECORD
    {
        uint ihFonts;
        LogFont elw = null;

        bool isExDv = false;

        internal EMR_EXTCREATEFONTINDIRECTW(LogFontExDv elw)
        {
            Type = RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW;
            Size = 368;
            ihFonts = 2;
            this.elw = elw;
        }

        internal EMR_EXTCREATEFONTINDIRECTW(BinaryReader br, uint TypeValue) : base(br, TypeValue)
        {
            ihFonts = br.ReadUInt32();

            var sizeOfVariableObject = Size - 12;
            if (sizeOfVariableObject > 320)//Size of a LogFontPanose object
            {
                elw = new LogFontExDv(br);
                isExDv = true;
            }
            else if (sizeOfVariableObject == 320)
            {
                //Fixed length is LogFontPanose
                elw = new LogFontPanose(br);
            }
            else
            {
                throw new InvalidOperationException("Corrupt file. The 'elw' field of a EXTCREATEFONTINDIRECTW object cannot be smaller than 320 bytes");
            }
        }

        internal override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihFonts);
            elw.WriteBytes(bw);
        }
    }
}
