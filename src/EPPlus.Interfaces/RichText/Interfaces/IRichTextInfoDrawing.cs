using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.RichText;

namespace OfficeOpenXml.Interfaces.RichText.Interfaces
{
    /// <summary>
    /// See 'ExcelParagraphTextRunBase.cs' for all options this interface should support
    /// </summary>
    public interface IRichTextInfoDrawing : IRichTextInfoSimple
    {
        DrawingStrikeType DrawingStrike{ get; set; } /*{ get { return (DrawingStrikeType)StrikeType; } set { StrikeType = (int)value; } }*/
        DrawingUnderlineStyle UnderlineStyle { get; set; }
        DrawingTextCapsType Capitalization { get; set; }
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
