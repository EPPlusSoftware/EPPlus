using OfficeOpenXml.Interfaces.RichText.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    public class RichTextDefaults : IRichTextInfoBase
    {
        internal RichTextDefaults()
        {
        }
        public bool IsItalic { get; set; } = false;

        public bool IsBold { get; set; } = false;

        public bool SubScript { get; set; } = false;

        public bool SuperScript { get; set; } = false;

        public int UnderlineType { get; set; } = -1;

        public int StrikeType { get; set; } = -1;

        public int Capitalization { get; set; } = -1;

        public Color UnderlineColor { get; set; }

        public Color FontColor { get; set; }

        //TODO Offset which is equal to 30% or -25% if Sub or Superscript are true?
    }
}
