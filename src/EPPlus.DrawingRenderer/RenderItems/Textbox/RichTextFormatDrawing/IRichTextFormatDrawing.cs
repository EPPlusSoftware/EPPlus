using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using System.Drawing;

namespace EPPlus.DrawingRenderer.RenderItems.Textbox
{
    /// <summary>
    /// TODO: Move this to interfaces. Only here in order to not break existing references in PDF 
    /// (This should be moved when IRichTextFormatSimple is moved)
    /// 
    /// Rich text data for drawings
    /// </summary>
    public interface IRichTextFormatDrawing : IRichTextFormatSimple
    {
        public new eDrawingStrikeType StrikeType { get; set; } /*{ get { return (DrawingStrikeType)StrikeType; } set { StrikeType = (int)value; } }*/
        public new eDrawingUnderLineType UnderlineType { get; set; }
        Color? HighLightColor { get; set; }
        /// <summary>
        /// The spacing between characters within a text run.
        /// </summary>
        double Spacing { get; set; }

        /// <summary>
        /// +Superscript or -Subscript offset in percent 
        /// (default 30% Super and -25% subscript)  
        /// </summary>
        public double Baseline { get; set; }

        //TODO: Advanced fills/Textoutline
        //TODO: Effects once implemented in Epplus
    }
}
