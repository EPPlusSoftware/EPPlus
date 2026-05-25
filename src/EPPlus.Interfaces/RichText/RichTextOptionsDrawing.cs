using OfficeOpenXml.Interfaces.Drawing.RichText;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace OfficeOpenXml.Interfaces.RichText
{
    internal class RichTextOptionsDrawing : IRichTextInfoDrawing
    {
        double _baseLine = 0d;
        public bool _subScript = false;
        public bool _superScript = false;

        public double Spacing { get; set; } = 0d;
        /// <summary>
        /// +Superscript or -Subscript offset in percent 
        /// (default 30% Super and -25% subscript)  
        /// </summary>
        public double Baseline
        {
            get
            {
                return _baseLine;
            }
            set
            {
                if (value > 0d)
                {
                    _superScript = true;
                    _subScript = false;
                }
                else if (value < 0d)
                {
                    _superScript = false;
                    _subScript = true;
                }
                else
                {
                    //When offset is 0 it is neither a sub or super script
                    _superScript = false;
                    _subScript = false;
                }

                _baseLine = value;
            }
        }
        public bool SubScript
        {
            get
            {
                return _subScript;
            }
            set
            {
                if (value == true)
                {
                    _superScript = false;
                    Baseline = -25d;
                }
                _subScript = value;
            }
        }

        public bool SuperScript
        {
            get
            {
                return _superScript;
            }
            set
            {
                if (value == true)
                {
                    _subScript = false;
                    Baseline = 30d;
                }
                _superScript = value;
            }
        }
        //---
        //StrikeType and UnderlineType being public makes this clumsy and confusing... Consider refactor or wrapping
        public DrawingStrikeType DrawingStrike { get { return (DrawingStrikeType)StrikeType; } set { StrikeType = (int)value; } }
        public DrawingUnderlineStyle UnderlineStyle { get { return (DrawingUnderlineStyle)StrikeType; } set { StrikeType = (int)value; } }
        ///---
        ///
        public DrawingTextCapsType Capitalization { get; set; } = DrawingTextCapsType.None;
        public Color? HighLightColor { get; set; } = null;
        public bool Italic { get; set; } = false;
        public bool Bold { get; set; } = false;
        public int UnderlineType { get; set; } = (int)DrawingUnderlineStyle.None;
        public int StrikeType { get; set; } = (int)DrawingStrikeType.No;
        public Color UnderlineColor { get; set; }
        public Color FontColor { get; set; }
        public string FontFamily { get; set; }
        public FontSubFamily SubFamily { get; set; }
        public double Size { get; set; }

        public void SetFont(IFontData font)
        {
            FontFamily = font.FontFamily;
            Size = font.Size;
            SubFamily = font.SubFamily;
        }
    }
}
