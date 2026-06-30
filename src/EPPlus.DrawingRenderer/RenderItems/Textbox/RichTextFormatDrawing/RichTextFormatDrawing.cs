using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;

namespace EPPlus.DrawingRenderer.RenderItems.Textbox
{
    /// <summary>
    /// Default richText data class for drawings
    /// </summary>
    public class RichTextFormatDrawing : RichTextFormatSimple, IRichTextFormatDrawing
    {
        public new eDrawingStrikeType StrikeType { get => (eDrawingStrikeType)base.StrikeType; set => base.StrikeType = (int)value; } 
        public new eDrawingUnderLineType UnderlineType { get => (eDrawingUnderLineType)base.UnderlineType; set => base.UnderlineType = (int)value; }

        public Color? HighLightColor { get; set; }
        public double Spacing { get; set; } = 0d;

        double _baseLine = 0d;

        private bool _subScript = false;
        private bool _superScript = false;

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
        public new bool SubScript
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

        public new bool SuperScript
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

        public RichTextFormatDrawing() : base()
        {
            StrikeType = eDrawingStrikeType.No;
            UnderlineType = eDrawingUnderLineType.None;
        }

        public RichTextFormatDrawing(string text, string fontFamily, float size, bool bold = false, bool italic = false) : base(text, fontFamily, size, bold, italic)
        {
            StrikeType = eDrawingStrikeType.No;
            UnderlineType = eDrawingUnderLineType.None;
        }

        public RichTextFormatDrawing(FontFormatBase defaultParagraphFont)
        {
            SetFont(defaultParagraphFont);
        }
    }
}
