using OfficeOpenXml.Utils;
using System.IO;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Globalization;
using System;

namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFont
    {
        //Read and written properties
        int height;
        internal int Width;
        internal int Escapement;
        internal int Orientation;
        int weight;
        internal byte Italic;
        internal byte Underline;
        internal byte StrikeOut;
        internal CharacterSet Set;
        internal byte OutPrecision;
        internal byte ClipPrecision;
        internal byte Quality;
        internal byte PitchAndFamily;
        string faceName;

        internal int Height
        {
            get
            {
                return height;
            }
            set
            {
                height = value;
                mFont.Size = FontPointSize;
            }
        }

        internal int Weight
        {
            get
            {
                return weight;
            }
            set
            {
                weight = value;
                var fontStyle = MeasurementFontStyles.Regular;
                if (Weight >= 700)
                {
                    fontStyle = MeasurementFontStyles.Bold;
                }
                mFont.Style = fontStyle;
            }
        }

        internal string FaceName
        {
            get
            {
                return faceName;
            }
            set
            {
                faceName = value;

                //Ensure faceName can be used by enums for fonts later
                TextInfo textInfo = CultureInfo.InvariantCulture.TextInfo;
                var changedFaceName = textInfo.ToTitleCase(FaceName);

                if (changedFaceName.Contains("Ui"))
                {
                    changedFaceName = changedFaceName.Replace("Ui", "UI");
                }

                mFont.FontFamily = changedFaceName;
            }
        }

        //Simplified properties for viewing/editing
        internal FamilyFont fontFamily;
        internal Pitch pitchFont;

        internal MeasurementFont mFont = new MeasurementFont();

        internal int FontPointSize
        {
            get
            {
                if (0 < Height)
                {
                    //TODO: Transform into device units
                    return Height;
                }
                if (Height == 0)
                {
                    return 11;
                }
                else
                {
                    return Height < 0 ? Math.Abs(Height) : Height;
                }
            }
        }

        internal int CalculatedAverageWidth;

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

            //Assuming output pixel width is equal to height
            CalculatedAverageWidth = Width != 0 ? Width : (int)Math.Round((FontPointSize / 2d),MidpointRounding.AwayFromZero);
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
