using System;
using System.IO;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class EMR_EXTCREATEFONTINDIRECTW : EMR_RECORD
    {
        internal uint ihFonts;
        internal LogFont elw = null;
        ExcelFont excelFont;

        bool isExDv = false;

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
            else if(sizeOfVariableObject == 92)
            {
                //The object MAY be as simple as a logfont object.
                elw = new LogFont(br);
            }
            else
            {
                throw new InvalidOperationException("Corrupt file. The 'elw' field of a EXTCREATEFONTINDIRECTW object cannot be smaller than 320 bytes");
            }
            //excelFont.Family = (int)elw.PitchAndFamily;
            //excelFont.

            //font.Style = MeasurementFontStyles.Regular;
            //font.FontFamily = elw.FaceName;
            //font.Size = 
        }

        public override void WriteBytes(BinaryWriter bw)
        {
            base.WriteBytes(bw);
            bw.Write(ihFonts);
            elw.WriteBytes(bw);
        }
    }
}
