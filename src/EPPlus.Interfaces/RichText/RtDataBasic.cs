using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class RtDataBasic : FontDataBasic
    {
        public RtDataBasic()
        {
        }

        public RtDataBasic(MeasurementFont font) : base(font)
        {
        }

        public RtDataBasic(string text, string fontFamily, FontSubFamily subFamily, double fontSize) : base(fontFamily, subFamily, fontSize)
        {
            Text = text;
        }

        public RtDataBasic(string text, string fontFamily, double fontSize) : base(fontFamily, FontSubFamily.Regular, fontSize)
        {
            Text = text;
        }

        public string Text { get; set; }

        public bool Italic
        {
            get { return (SubFamily & FontSubFamily.Italic) == FontSubFamily.Italic; }
            set
            {
                if (value)
                {
                    //Set Flag
                    SubFamily = SubFamily | FontSubFamily.Italic;

                }
                else
                {
                    //Unset Flag
                    SubFamily &= ~FontSubFamily.Italic;
                }
            }
        }

        public bool Bold
        {
            get { return (SubFamily & FontSubFamily.Bold) == FontSubFamily.Bold; }
            set
            {
                if (value)
                {
                    //Set flag
                    SubFamily = SubFamily | FontSubFamily.Bold;

                }
                else
                {
                    //Unset flag
                    SubFamily &= ~FontSubFamily.Bold;
                }
            }
        }

        public double AscentPoints { get; set; }
        public double DescentPoints { get; set; }
    }
}
