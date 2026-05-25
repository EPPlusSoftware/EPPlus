using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// Holds the basic data required for the textmeasurer
    /// A rich text object may contain more data than this class
    /// but a richText object MUST inculde At Least this much.
    /// </summary>
    public class RtDataBasic : FontDataBasic, IRichTextInfoEssential
    {
        public RtDataBasic()
        {
        }

        /// <summary>
        /// Legacy constructor. Prefer to avoid with new implementations. To be removed after refactor
        /// </summary>
        /// <param name="font"></param>
        public RtDataBasic(MeasurementFont font) : base(font)
        {
        }

        public RtDataBasic(string text, string fontFamily, FontSubFamily subFamily, float fontSize) : base(fontFamily, subFamily, fontSize)
        {
            Text = text;
        }

        public RtDataBasic(string text, string fontFamily, float fontSize) : base(fontFamily, FontSubFamily.Regular, fontSize)
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
