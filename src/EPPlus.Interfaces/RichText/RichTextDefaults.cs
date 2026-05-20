using OfficeOpenXml.Interfaces.RichText.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using OfficeOpenXml.Interfaces.Fonts;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// Basic default richTextData including default font
    /// </summary>
    public class RichTextDefaults : IRichTextInfoBase
    {
        public RichTextDefaults()
        {
        }

        public bool Italic 
        { 
            get { return (SubFamily & FontSubFamily.Italic) == FontSubFamily.Italic; } 
            set 
            {
                if(value)
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
        public bool SubScript { get; set; } = false;

        public bool SuperScript { get; set; } = false;

        public int UnderlineType { get; set; } = -1;

        public int StrikeType { get; set; } = -1;

        public int Capitalization { get; set; } = -1;

        public Color UnderlineColor { get; set; }

        public Color FontColor { get; set; }

        public string FontFamily { get; set; } = "Archivo Narrow";
        public double FontSize { get; set; } = 11d;
        public FontSubFamily SubFamily { get; set; } = FontSubFamily.Regular;

        //TODO Offset which is equal to 30% or -25% if Sub or Superscript are true?
    }
}
