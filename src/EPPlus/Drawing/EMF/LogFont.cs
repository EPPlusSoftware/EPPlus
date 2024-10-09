using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Collections;


namespace OfficeOpenXml.Drawing.EMF
{
    internal class LogFont
    {
        //Read and written properties
        internal int Height;
        int Width;
        int Escapement;
        int Orientation;
        int Weight;
        byte Italic;
        byte Underline;
        byte StrikeOut;
        CharacterSet Set;
        byte OutPrecision;
        byte ClipPrecision;
        byte Quality;
        internal byte PitchAndFamily;
        internal string FaceName;

        internal MeasurementFont mFont = new MeasurementFont();

        internal int FontPointSize
        {
            get
            {
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

        //Simplified properties for viewing/editing
        FamilyFont fontFamily;
        Pitch pitchFont;

        internal int CalculatedAverageWidth;
        internal int DefinedHeight;
        //internal float OneDesignUnit;

        private bool recalculateWidth = false;

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

            DefinedHeight = DefineHeight();
            //Assuming output pixel width is equal to height
            CalculatedAverageWidth = Width != 0 ? Width : (int)Math.Round((DefinedHeight / 2d),MidpointRounding.AwayFromZero);

            TextInfo textInfo = CultureInfo.InvariantCulture.TextInfo;
            var changedFaceName = textInfo.ToTitleCase(FaceName);

            if(changedFaceName.Contains("Ui"))
            {
                changedFaceName = changedFaceName.Replace("Ui", "UI");
            }


            //Thin 100
            //Extra Light(Ultra Light) 200
            //Light 300
            //Normal(Regular) 400
            //Medium 500
            //Semi - Bold(Demi - Bold) 600
            //Bold 700
            //Extra Bold(Ultra Bold) 800
            //Heavy(Black) 900

            var fontStyle = MeasurementFontStyles.Regular;
            if(Weight >= 700)
            {
                fontStyle = MeasurementFontStyles.Bold;
            }

            mFont = new MeasurementFont()
            {
                FontFamily = changedFaceName,
                Size = FontPointSize,
                Style = fontStyle
            };
        }

        //internal void CalculateUnitsPerEm(float ppi)
        //{
        //    var height = DefineHeight();
        //    float heightInPt = height * 0.75f;
        //    OneDesignUnit = ((heightInPt / 72f) * ppi) / height;
        //}

        int DefineHeight()
        {
            if(0 < Height)
            {
                //TODO: Transform into device units
                return Height;
            }
            else if(Height == 0)
            {
                return 11;
            }
            else 
            {
                return Math.Abs(Height);
            }
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
